#include "AutoLinkerTestApi.h"

#include <algorithm>
#include <atomic>
#include <chrono>
#include <cstring>
#include <filesystem>
#include <fstream>
#include <format>
#include <string>
#include <thread>
#include <vector>
#include <winsock2.h>
#include <ws2tcpip.h>

#include "..\\thirdparty\\json.hpp"

#include "AIChatTooling.h"
#include "AIChatMcpClient.h"
#include "AIChatMcpConfig.h"
#include "AIChatToolPolicy.h"
#include "AIService.h"
#include "AutoLinkerVersion.h"
#include "GameAnalyticsClient.h"
#include "LocalMcpServer.h"
#include "PathHelper.h"
#include "PowerShellToolRunner.h"
#include "RealPageCodeToolSupport.h"
#include "Version.h"
#include "WebDocumentClient.h"
#include "WebDocumentExtractor.h"
#include "WorkspaceFileTools.h"
#include "WorkspaceMirror.h"

#pragma comment(lib, "Ws2_32.lib")

namespace {

constexpr const char* kUtf8ChineseArgument = "\xE4\xB8\xAD\xE6\x96\x87\xE5\x8F\x82\xE6\x95\xB0";
constexpr const char* kUtf8TestPassed = "\xE6\xB5\x8B\xE8\xAF\x95\xE9\x80\x9A\xE8\xBF\x87";

std::string LocalToUtf8ForTest(const std::string& text)
{
	if (text.empty()) {
		return std::string();
	}
	const int wideLength = MultiByteToWideChar(
		CP_ACP,
		0,
		text.data(),
		static_cast<int>(text.size()),
		nullptr,
		0);
	if (wideLength <= 0) {
		return text;
	}
	std::wstring wide(static_cast<size_t>(wideLength), L'\0');
	if (MultiByteToWideChar(CP_ACP, 0, text.data(), static_cast<int>(text.size()), wide.data(), wideLength) <= 0) {
		return text;
	}
	const int utf8Length = WideCharToMultiByte(
		CP_UTF8,
		0,
		wide.data(),
		wideLength,
		nullptr,
		0,
		nullptr,
		nullptr);
	if (utf8Length <= 0) {
		return text;
	}
	std::string utf8(static_cast<size_t>(utf8Length), '\0');
	if (WideCharToMultiByte(
			CP_UTF8,
			0,
			wide.data(),
			wideLength,
			utf8.data(),
			utf8Length,
			nullptr,
			nullptr) <= 0) {
		return text;
	}
	return utf8;
}

bool RunAIChatToolPolicySelfTest(nlohmann::json& outCheck)
{
	outCheck = {
		{"name", "ai_chat_tool_policy"},
		{"ok", false}
	};

	AIChatToolPolicy::Session duplicatePolicy;
	const std::string readArgs = R"({"file_path":"src/Test.txt","offset":0,"limit":10})";
	const auto firstRead = duplicatePolicy.BeforeToolCall("read_file", readArgs);
	duplicatePolicy.AfterToolCall(
		"read_file",
		readArgs,
		R"({"ok":true,"file_path":"src/Test.txt","offset":0,"returned_lines":10,"content":"1\ta\n"})",
		true);
	const auto duplicateRead = duplicatePolicy.BeforeToolCall("read_file", readArgs);

	AIChatToolPolicy::Session batchCoveragePolicy;
	const std::string batchReadArgs =
		R"({"files":[{"file_path":"src/Batch.txt","offset":0,"limit":10}]})";
	const auto firstBatchRead = batchCoveragePolicy.BeforeToolCall("read_files", batchReadArgs);
	batchCoveragePolicy.AfterToolCall(
		"read_files",
		batchReadArgs,
		R"({"ok":true,"files":[{"ok":true,"file_path":"src/Batch.txt","offset":0,"returned_lines":10,"total_lines":30,"total_lines_complete":true,"content":"1\ta\n10\tz\n"}]})",
		true);
	const auto overlappingBatchRead = batchCoveragePolicy.BeforeToolCall(
		"read_files",
		R"({"files":[{"file_path":"src/Batch.txt","offset":5,"limit":15}]})");
	const nlohmann::json overlapResult = nlohmann::json::parse(
		overlappingBatchRead.resultJsonUtf8,
		nullptr,
		false);
	const bool overlapSuggestionOk =
		!overlappingBatchRead.allowed &&
		overlappingBatchRead.reason == "overlapping_read_range" &&
		overlapResult.is_object() &&
		overlapResult.contains("suggested_missing_ranges") &&
		overlapResult["suggested_missing_ranges"].is_array() &&
		overlapResult["suggested_missing_ranges"].size() == 1 &&
		overlapResult["suggested_missing_ranges"][0].value("offset", -1) == 10 &&
		overlapResult["suggested_missing_ranges"][0].value("limit", -1) == 10;

	AIChatToolPolicy::Session budgetPolicy;
	bool budgetCallsAllowed = true;
	for (int i = 0; i < AIChatToolPolicy::kHardExplorationCallLimit; ++i) {
		const std::string args = std::format(R"({{"pattern":"item{}","regex":false}})", i);
		const auto decision = budgetPolicy.BeforeToolCall("search_code", args);
		budgetCallsAllowed = budgetCallsAllowed && decision.allowed;
		budgetPolicy.AfterToolCall("search_code", args, R"({"ok":true,"results":[]})", true);
	}
	const auto overBudget = budgetPolicy.BeforeToolCall(
		"search_code",
		R"({"pattern":"over-budget","regex":false})");

	AIChatToolPolicy::Session writePolicy;
	const std::string writeArgs = R"({"file_path":"src/Test.txt","full_code":"x"})";
	const auto writeDecision = writePolicy.BeforeToolCall("write_file", writeArgs);
	const std::string writeNotice = writePolicy.AfterToolCall(
		"write_file",
		writeArgs,
		R"({"ok":true,"verified":true,"new_hash":"abc"})",
		true);
	const bool prefersLowThinking = writePolicy.PreferLowThinkingForNextRound();
	const auto postWriteRead = writePolicy.BeforeToolCall(
		"read_files",
		R"({"file_paths":["src/Test.txt"]})");

	const bool ok = firstRead.allowed &&
		!duplicateRead.allowed &&
		firstBatchRead.allowed &&
		overlapSuggestionOk &&
		budgetCallsAllowed &&
		!overBudget.allowed &&
		writeDecision.allowed &&
		!writeNotice.empty() &&
		prefersLowThinking &&
		!postWriteRead.allowed;
	outCheck["ok"] = ok;
	outCheck["duplicate_reason"] = duplicateRead.reason;
	outCheck["batch_overlap_reason"] = overlappingBatchRead.reason;
	outCheck["batch_overlap_result"] = overlapResult;
	outCheck["budget_reason"] = overBudget.reason;
	outCheck["post_write_reason"] = postWriteRead.reason;
	return ok;
}

int CopyStringToBuffer(const std::string& value, char* buffer, int bufferSize)
{
	if (buffer == nullptr || bufferSize <= 0) {
		return AUTOLINKER_TEST_STRING_INVALID_ARGUMENT;
	}

	const size_t requiredSize = value.size() + 1;
	if (requiredSize > static_cast<size_t>(bufferSize)) {
		return AUTOLINKER_TEST_STRING_BUFFER_TOO_SMALL;
	}

	std::memcpy(buffer, value.c_str(), requiredSize);
	return static_cast<int>(value.size());
}

std::string BuildDeepSeekToolArgumentsJson(const std::string& toolName)
{
	if (toolName == "fetch_url") {
		return R"({"url":"https://api-docs.deepseek.com/quick_start/rate_limit","timeout_seconds":30,"max_bytes":262144})";
	}
	if (toolName == "extract_web_document") {
		return R"({"url":"https://api-docs.deepseek.com/guides/thinking_mode","timeout_seconds":30,"max_bytes":262144})";
	}
	return "{}";
}

nlohmann::json BuildDeepSeekIntegrationResultJson(const AISettings& settings)
{
	return {
		{"provider", "deepseek"},
		{"model", settings.model},
		{"base_url", settings.baseUrl},
		{"thinking_level", AIService::ThinkingLevelToString(settings.thinkingLevel)},
		{"protocol", AIService::ProtocolTypeToString(settings.protocolType)}
	};
}

std::string BuildOpenAIToolArgumentsJson(const std::string& toolName)
{
	if (toolName == "fetch_url") {
		return R"({"url":"https://developers.openai.com/api/docs/api-reference/chat/create-chat-completion","timeout_seconds":30,"max_bytes":262144})";
	}
	if (toolName == "extract_web_document") {
		return R"({"url":"https://developers.openai.com/api/docs/api-reference/responses/create","timeout_seconds":30,"max_bytes":262144})";
	}
	return "{}";
}

nlohmann::json BuildOpenAIIntegrationResultJson(const AISettings& settings)
{
	return {
		{"provider", "openai"},
		{"model", settings.model},
		{"base_url", settings.baseUrl},
		{"thinking_level", AIService::ThinkingLevelToString(settings.thinkingLevel)},
		{"protocol", AIService::ProtocolTypeToString(settings.protocolType)}
	};
}

std::string BuildGeminiToolArgumentsJson(const std::string& toolName)
{
	if (toolName == "fetch_url") {
		return R"({"url":"https://ai.google.dev/gemini-api/docs/models/gemini","timeout_seconds":30,"max_bytes":4096})";
	}
	if (toolName == "extract_web_document") {
		return R"({"url":"https://ai.google.dev/gemini-api/docs/function-calling","timeout_seconds":30,"max_bytes":4096})";
	}
	return "{}";
}

nlohmann::json BuildGeminiIntegrationResultJson(const AISettings& settings)
{
	return {
		{"provider", "gemini"},
		{"model", settings.model},
		{"base_url", settings.baseUrl},
		{"thinking_level", AIService::ThinkingLevelToString(settings.thinkingLevel)},
		{"protocol", AIService::ProtocolTypeToString(settings.protocolType)}
	};
}

std::string BuildClaudeToolArgumentsJson(const std::string& toolName)
{
	if (toolName == "fetch_url") {
		return R"({"url":"https://docs.anthropic.com/en/docs/about-claude/models/overview","timeout_seconds":30,"max_bytes":4096})";
	}
	if (toolName == "extract_web_document") {
		return R"({"url":"https://docs.anthropic.com/en/docs/agents-and-tools/tool-use/overview","timeout_seconds":30,"max_bytes":4096})";
	}
	return "{}";
}

nlohmann::json BuildClaudeIntegrationResultJson(const AISettings& settings)
{
	return {
		{"provider", "claude"},
		{"model", settings.model},
		{"base_url", settings.baseUrl},
		{"thinking_level", AIService::ThinkingLevelToString(settings.thinkingLevel)},
		{"protocol", AIService::ProtocolTypeToString(settings.protocolType)}
	};
}

std::string ExtractMessageContentText(const nlohmann::json& parsed)
{
	if (!parsed.is_object() || !parsed.contains("content")) {
		return std::string();
	}
	if (parsed["content"].is_string()) {
		return parsed["content"].get<std::string>();
	}
	if (!parsed["content"].is_array()) {
		return std::string();
	}

	std::string content;
	for (const auto& item : parsed["content"]) {
		if (!item.is_object()) {
			continue;
		}
		const std::string contentType = item.value("type", std::string());
		if ((contentType == "output_text" || contentType == "text") &&
			item.contains("text") &&
			item["text"].is_string()) {
			content += item["text"].get<std::string>();
		}
	}
	return content;
}

std::vector<AIChatMessage> BuildFollowupMessagesFromChatResult(
	const std::vector<AIChatMessage>& prefixMessages,
	const AIChatResult& toolChatResult)
{
	std::vector<AIChatMessage> followupMessages = prefixMessages;
	for (const auto& rawMessageJsonUtf8 : toolChatResult.contextPrefixRawMessagesUtf8) {
		nlohmann::json parsed;
		try {
			parsed = nlohmann::json::parse(rawMessageJsonUtf8);
		}
		catch (...) {
			continue;
		}
		if (!parsed.is_object()) {
			continue;
		}

		const std::string role = parsed.value("role", std::string());
		const std::string type = parsed.value("type", std::string());
		if (role == "assistant") {
			followupMessages.push_back({
				"assistant",
				ExtractMessageContentText(parsed),
				parsed.value("reasoning_content", std::string()),
				rawMessageJsonUtf8
			});
		}
		else if (role == "tool") {
			followupMessages.push_back({
				"tool",
				ExtractMessageContentText(parsed),
				"",
				rawMessageJsonUtf8
			});
		}
		else if (!type.empty()) {
			followupMessages.push_back({
				"tool",
				ExtractMessageContentText(parsed),
				"",
				rawMessageJsonUtf8
			});
		}
	}
	return followupMessages;
}

std::string DumpJsonPrettySafe(const nlohmann::json& value)
{
	return value.dump(2, ' ', false, nlohmann::json::error_handler_t::replace);
}

std::string ReadFileBinary(const std::filesystem::path& path)
{
	std::ifstream input(path, std::ios::binary);
	if (!input.is_open()) {
		return std::string();
	}
	return std::string(std::istreambuf_iterator<char>(input), std::istreambuf_iterator<char>());
}

bool WriteFileBinary(const std::filesystem::path& path, const std::string& text)
{
	std::error_code ec;
	const auto parent = path.parent_path();
	if (!parent.empty()) {
		std::filesystem::create_directories(parent, ec);
	}
	std::ofstream output(path, std::ios::binary | std::ios::trunc);
	if (!output.is_open()) {
		return false;
	}
	output.write(text.data(), static_cast<std::streamsize>(text.size()));
	return output.good();
}

class MockMcpHttpServer {
public:
	~MockMcpHttpServer()
	{
		Stop();
	}

	bool Start(std::string& outError)
	{
		outError.clear();
		WSADATA wsa = {};
		if (WSAStartup(MAKEWORD(2, 2), &wsa) != 0) {
			outError = "WSAStartup failed";
			return false;
		}
		wsaStarted_ = true;

		listenSocket_ = socket(AF_INET, SOCK_STREAM, IPPROTO_TCP);
		if (listenSocket_ == INVALID_SOCKET) {
			outError = "socket failed";
			return false;
		}

		sockaddr_in addr = {};
		addr.sin_family = AF_INET;
		addr.sin_addr.s_addr = htonl(INADDR_LOOPBACK);
		addr.sin_port = 0;
		if (bind(listenSocket_, reinterpret_cast<sockaddr*>(&addr), sizeof(addr)) == SOCKET_ERROR) {
			outError = "bind failed";
			return false;
		}
		if (listen(listenSocket_, SOMAXCONN) == SOCKET_ERROR) {
			outError = "listen failed";
			return false;
		}

		sockaddr_in bound = {};
		int boundLen = sizeof(bound);
		if (getsockname(listenSocket_, reinterpret_cast<sockaddr*>(&bound), &boundLen) == SOCKET_ERROR) {
			outError = "getsockname failed";
			return false;
		}
		port_ = ntohs(bound.sin_port);
		stop_.store(false);
		thread_ = std::thread([this]() { Run(); });
		return true;
	}

	void Stop()
	{
		stop_.store(true);
		if (listenSocket_ != INVALID_SOCKET) {
			closesocket(listenSocket_);
			listenSocket_ = INVALID_SOCKET;
		}
		if (thread_.joinable()) {
			thread_.join();
		}
		if (wsaStarted_) {
			WSACleanup();
			wsaStarted_ = false;
		}
	}

	std::string Url() const
	{
		return "http://127.0.0.1:" + std::to_string(port_) + "/mcp";
	}

private:
	void Run()
	{
		while (!stop_.load()) {
			fd_set readSet = {};
			FD_ZERO(&readSet);
			FD_SET(listenSocket_, &readSet);
			timeval timeout = {};
			timeout.tv_sec = 0;
			timeout.tv_usec = 200000;
			const int selected = select(0, &readSet, nullptr, nullptr, &timeout);
			if (selected <= 0 || stop_.load()) {
				continue;
			}
			SOCKET client = accept(listenSocket_, nullptr, nullptr);
			if (client == INVALID_SOCKET) {
				continue;
			}
			HandleClient(client);
			closesocket(client);
		}
	}

	static std::string BuildHttpResponse(int status, const std::string& body, bool includeSessionHeader)
	{
		const char* statusText = status == 202 ? "Accepted" : "OK";
		std::string response = "HTTP/1.1 " + std::to_string(status) + " " + statusText + "\r\n";
		response += "Connection: close\r\n";
		response += "Content-Type: application/json\r\n";
		if (includeSessionHeader) {
			response += "Mcp-Session-Id: mock-session\r\n";
		}
		response += "Content-Length: " + std::to_string(body.size()) + "\r\n\r\n";
		response += body;
		return response;
	}

	static int ParseContentLength(const std::string& headers)
	{
		const std::string marker = "Content-Length:";
		size_t pos = headers.find(marker);
		if (pos == std::string::npos) {
			pos = headers.find("content-length:");
		}
		if (pos == std::string::npos) {
			return 0;
		}
		pos += marker.size();
		while (pos < headers.size() && (headers[pos] == ' ' || headers[pos] == '\t')) {
			++pos;
		}
		size_t end = pos;
		while (end < headers.size() && headers[end] >= '0' && headers[end] <= '9') {
			++end;
		}
		try {
			return std::stoi(headers.substr(pos, end - pos));
		}
		catch (...) {
			return 0;
		}
	}

	static std::string BuildRpcBody(const nlohmann::json& request)
	{
		const std::string method = request.value("method", std::string());
		nlohmann::json response = {
			{"jsonrpc", "2.0"}
		};
		if (request.contains("id")) {
			response["id"] = request["id"];
		}

		if (method == "initialize") {
			response["result"] = {
				{"protocolVersion", "2025-11-25"},
				{"capabilities", {{"tools", nlohmann::json::object()}}},
				{"serverInfo", {{"name", "AutoLinkerTest Mock MCP"}, {"version", "1.0"}}}
			};
		}
		else if (method == "tools/list") {
			response["result"] = {
				{"tools", nlohmann::json::array({
					{
						{"name", "echo"},
						{"description", "Echo UTF-8 text for AutoLinker MCP tests."},
						{"inputSchema", {
							{"type", "object"},
							{"properties", {
								{"text", {{"type", "string"}, {"description", "Text to echo."}}}
							}},
							{"required", nlohmann::json::array({"text"})},
							{"additionalProperties", false}
						}}
					}
				})}
			};
		}
		else if (method == "tools/call") {
			const nlohmann::json params = request.value("params", nlohmann::json::object());
			const nlohmann::json args = params.value("arguments", nlohmann::json::object());
			const std::string text = args.value("text", std::string());
			response["result"] = {
				{"content", nlohmann::json::array({
					{{"type", "text"}, {"text", "mock echo: " + text}}
				})},
				{"structuredContent", {{"echo", text}, {"utf8", kUtf8TestPassed}}},
				{"isError", false}
			};
		}
		else {
			response["result"] = nlohmann::json::object();
		}
		return response.dump(-1, ' ', false, nlohmann::json::error_handler_t::replace);
	}

	void HandleClient(SOCKET client)
	{
		std::string requestText;
		char buffer[4096] = {};
		while (requestText.find("\r\n\r\n") == std::string::npos) {
			const int received = recv(client, buffer, sizeof(buffer), 0);
			if (received <= 0) {
				return;
			}
			requestText.append(buffer, static_cast<size_t>(received));
			if (requestText.size() > 1024 * 1024) {
				return;
			}
		}
		const size_t headerEnd = requestText.find("\r\n\r\n");
		const std::string headers = requestText.substr(0, headerEnd);
		const int contentLength = ParseContentLength(headers);
		while (static_cast<int>(requestText.size() - headerEnd - 4) < contentLength) {
			const int received = recv(client, buffer, sizeof(buffer), 0);
			if (received <= 0) {
				return;
			}
			requestText.append(buffer, static_cast<size_t>(received));
		}
		const std::string body = requestText.substr(headerEnd + 4, static_cast<size_t>(contentLength));
		const nlohmann::json request = nlohmann::json::parse(body, nullptr, false);
		if (request.is_discarded()) {
			const std::string response = BuildHttpResponse(200, R"({"jsonrpc":"2.0","error":{"code":-32700,"message":"parse error"}})", true);
			send(client, response.data(), static_cast<int>(response.size()), 0);
			return;
		}
		if (request.value("method", std::string()) == "notifications/initialized") {
			const std::string response = BuildHttpResponse(202, "", true);
			send(client, response.data(), static_cast<int>(response.size()), 0);
			return;
		}
		const std::string responseBody = BuildRpcBody(request);
		const std::string response = BuildHttpResponse(200, responseBody, true);
		send(client, response.data(), static_cast<int>(response.size()), 0);
	}

	bool wsaStarted_ = false;
	SOCKET listenSocket_ = INVALID_SOCKET;
	unsigned short port_ = 0;
	std::atomic_bool stop_{ false };
	std::thread thread_;
};

class ScopedMcpConfigBackup {
public:
	ScopedMcpConfigBackup()
		: path_(AIChatMcpConfigStore::GetConfigPath()),
		  hadConfig_(std::filesystem::exists(path_)),
		  backup_(hadConfig_ ? ReadFileBinary(path_) : std::string())
	{
	}

	~ScopedMcpConfigBackup()
	{
		std::error_code ec;
		if (hadConfig_) {
			WriteFileBinary(path_, backup_);
		}
		else {
			std::filesystem::remove(path_, ec);
		}
	}

private:
	std::filesystem::path path_;
	bool hadConfig_ = false;
	std::string backup_;
};

bool RunMcpMockRoundtripSelfTest(nlohmann::json& outCheck)
{
	outCheck = {
		{"name", "mock_streamable_http_roundtrip"},
		{"ok", false}
	};

	ScopedMcpConfigBackup configBackup;
	MockMcpHttpServer server;
	std::string error;
	try {
		if (!server.Start(error)) {
			outCheck["error"] = error;
			return false;
		}

		AIChatMcpConfig config;
		AIChatMcpServerConfig mcpServer;
		mcpServer.id = "mock-mcp";
		mcpServer.name = "Mock MCP";
		mcpServer.transport = "streamable_http";
		mcpServer.url = server.Url();
		mcpServer.enabled = true;
		mcpServer.timeoutMs = 30000;
		config.servers.push_back(std::move(mcpServer));
		if (!AIChatMcpConfigStore::Save(config, &error)) {
			outCheck["error"] = error;
			return false;
		}

		const std::vector<AIChatMcpToolInfo> tools = AIChatMcpClient::LoadEnabledTools();
		outCheck["listed_tool_count"] = tools.size();
		auto it = std::find_if(tools.begin(), tools.end(), [](const AIChatMcpToolInfo& tool) {
			return tool.serverId == "mock-mcp" && tool.originalName == "echo";
		});
		if (it == tools.end()) {
			outCheck["error"] = "mock echo tool not listed";
			return false;
		}
		outCheck["model_tool_name"] = it->modelName;

		bool publicCatalogContainsOutboundMcp = false;
		const nlohmann::json publicCatalog = nlohmann::json::parse(
			AIService::BuildPublicToolCatalogJson(),
			nullptr,
			false);
		if (publicCatalog.is_array()) {
			for (const auto& item : publicCatalog) {
				if (item.is_object() && AIChatMcpClient::IsMcpModelToolName(item.value("name", std::string()))) {
					publicCatalogContainsOutboundMcp = true;
					break;
				}
			}
		}
		outCheck["public_catalog_contains_outbound_mcp"] = publicCatalogContainsOutboundMcp;

		bool approved = false;
		const AIChatMcpExecutionResult execution = AIChatMcpClient::ExecuteTool(
			it->modelName,
			R"({"text":"\u4e2d\u6587\u53c2\u6570"})",
			[&approved](const AIChatMcpApprovalContext&, bool& outAutoAllow, bool& outAutoAllowServer) {
				approved = true;
				outAutoAllow = false;
				outAutoAllowServer = true;
				return true;
			});
		AIChatMcpConfig savedConfig;
		const bool loadedSavedConfig = AIChatMcpConfigStore::Load(savedConfig, &error);
		// "Allow all methods" now snapshots per-tool grants (server id + tool name
		// + real schema hash) for every currently-known tool instead of a blanket
		// "*"/"*" wildcard, so new/changed tools still re-prompt. Verify the echo
		// tool's grant was recorded with its actual schema hash.
		const std::string grantedSchemaHash = it->schemaHash;
		const bool serverGrantSaved = loadedSavedConfig && std::any_of(
			savedConfig.approvalGrants.begin(),
			savedConfig.approvalGrants.end(),
			[&grantedSchemaHash](const AIChatMcpApprovalGrant& grant) {
				return grant.serverId == "mock-mcp" &&
					grant.toolName == "echo" &&
					grant.schemaHash == grantedSchemaHash;
			});
		bool secondApprovalCalled = false;
		const AIChatMcpExecutionResult secondExecution = AIChatMcpClient::ExecuteTool(
			it->modelName,
			R"({"text":"server-grant"})",
			[&secondApprovalCalled](const AIChatMcpApprovalContext&, bool& outAutoAllow, bool& outAutoAllowServer) {
				secondApprovalCalled = true;
				outAutoAllow = false;
				outAutoAllowServer = false;
				return false;
			});
		const std::string executionResultUtf8 = LocalToUtf8ForTest(execution.resultJsonLocal);
		const std::string secondResultUtf8 = LocalToUtf8ForTest(secondExecution.resultJsonLocal);
		const nlohmann::json executionJson = nlohmann::json::parse(executionResultUtf8, nullptr, false);
		const bool exactUtf8 = executionJson.is_object() &&
			executionJson.contains("structuredContent") &&
			executionJson["structuredContent"].is_object() &&
			executionJson["structuredContent"].value("echo", std::string()) == kUtf8ChineseArgument &&
			executionJson["structuredContent"].value("utf8", std::string()) == kUtf8TestPassed;
		outCheck["approval_called"] = approved;
		outCheck["execution_ok"] = execution.ok;
		outCheck["result"] = executionResultUtf8;
		outCheck["exact_utf8"] = exactUtf8;
		outCheck["server_grant_saved"] = serverGrantSaved;
		outCheck["second_approval_called"] = secondApprovalCalled;
		outCheck["second_execution_ok"] = secondExecution.ok;
		outCheck["second_result"] = secondResultUtf8;
		outCheck["ok"] = approved &&
			execution.ok &&
			executionResultUtf8.find("mock echo") != std::string::npos &&
			exactUtf8 &&
			!publicCatalogContainsOutboundMcp &&
			serverGrantSaved &&
			!secondApprovalCalled &&
			secondExecution.ok &&
			secondResultUtf8.find("mock echo") != std::string::npos;
		if (!outCheck.value("ok", false)) {
			if (!loadedSavedConfig) {
				outCheck["error"] = error.empty() ? "reload MCP config failed" : error;
			}
			else if (!execution.errorUtf8.empty()) {
				outCheck["error"] = execution.errorUtf8;
			}
			else if (!secondExecution.errorUtf8.empty()) {
				outCheck["error"] = secondExecution.errorUtf8;
			}
			else {
				outCheck["error"] = "mock tools/call or server approval grant failed";
			}
		}
		return outCheck.value("ok", false);
	}
	catch (const std::exception& ex) {
		outCheck["error"] = ex.what();
		return false;
	}
	catch (...) {
		outCheck["error"] = "unknown exception";
		return false;
	}
}

bool RunMcpStdioRoundtripSelfTest(nlohmann::json& outCheck)
{
	outCheck = {
		{"name", "mock_stdio_roundtrip"},
		{"ok", false}
	};

	ScopedMcpConfigBackup configBackup;
	std::string error;
	try {
		char modulePath[MAX_PATH] = {};
		if (GetModuleFileNameA(nullptr, modulePath, static_cast<DWORD>(sizeof(modulePath))) <= 0) {
			outCheck["error"] = "GetModuleFileNameA failed";
			return false;
		}

		AIChatMcpConfig config;
		AIChatMcpServerConfig server;
		server.id = "mock-stdio";
		server.name = "Mock Stdio";
		server.transport = "stdio";
		server.command = modulePath;
		server.arguments = { "mock-mcp-stdio" };
		server.enabled = true;
		server.timeoutMs = 30000;
		config.servers.push_back(std::move(server));
		if (!AIChatMcpConfigStore::Save(config, &error)) {
			outCheck["error"] = error;
			return false;
		}

		const std::vector<AIChatMcpToolInfo> tools = AIChatMcpClient::LoadEnabledTools();
		outCheck["listed_tool_count"] = tools.size();
		auto it = std::find_if(tools.begin(), tools.end(), [](const AIChatMcpToolInfo& tool) {
			return tool.serverId == "mock-stdio" && tool.originalName == "echo";
		});
		if (it == tools.end()) {
			outCheck["error"] = "mock stdio echo tool not listed";
			return false;
		}
		outCheck["model_tool_name"] = it->modelName;

		bool approved = false;
		const AIChatMcpExecutionResult execution = AIChatMcpClient::ExecuteTool(
			it->modelName,
			R"({"text":"stdio-ok"})",
			[&approved](const AIChatMcpApprovalContext&, bool& outAutoAllow, bool& outAutoAllowServer) {
				approved = true;
				outAutoAllow = false;
				outAutoAllowServer = false;
				return true;
			});
		outCheck["approval_called"] = approved;
		outCheck["execution_ok"] = execution.ok;
		outCheck["result"] = execution.resultJsonLocal;
		outCheck["ok"] = approved && execution.ok && execution.resultJsonLocal.find("mock echo") != std::string::npos;
		if (!outCheck.value("ok", false)) {
			outCheck["error"] = execution.errorUtf8.empty() ? "mock stdio tools/call failed" : execution.errorUtf8;
		}
		return outCheck.value("ok", false);
	}
	catch (const std::exception& ex) {
		outCheck["error"] = ex.what();
		return false;
	}
	catch (...) {
		outCheck["error"] = "unknown exception";
		return false;
	}
}

// Regression coverage for MCP config parsing invariants added during the MCP
// review: mcpServers preservation on save, CRLF header-injection rejection,
// duplicate-id de-duplication, and explicit-transport handling.
bool RunMcpConfigInvariantsSelfTest(nlohmann::json& outCheck)
{
	outCheck = { {"name", "config_invariants"}, {"ok", false} };
	std::vector<std::string> failures;
	ScopedMcpConfigBackup configBackup;

	// 1. mcpServers block survives a servers-only save (WebView-style edit), but a
	//    native raw-JSON save that drops it deletes it.
	{
		const std::string source = R"({
			"version": 1,
			"servers": [ {"id":"http-a","name":"A","transport":"streamable_http","url":"http://127.0.0.1:9/mcp","enabled":false} ],
			"mcpServers": { "extern": {"command":"npx","args":["-y","srv"]} }
		})";
		AIChatMcpConfig parsed;
		std::string error;
		if (!AIChatMcpConfigStore::ParseConfigJson(source, parsed, error)) {
			failures.push_back("parse mcpServers config failed: " + error);
		}
		else {
			const bool externReadOnly = std::any_of(parsed.servers.begin(), parsed.servers.end(),
				[](const AIChatMcpServerConfig& s) { return s.id == "extern" && s.readOnly && s.transport == "stdio"; });
			if (!externReadOnly) {
				failures.push_back("mcpServers entry not parsed as read-only stdio");
			}
		}

		// Seed disk with the full config, then emulate a WebView save (servers only)
		// with preserve=true: mcpServers must survive.
		const std::string uiSave = R"({"version":1,"servers":[{"id":"http-a","name":"A","transport":"streamable_http","url":"http://127.0.0.1:9/mcp","enabled":false}],"approval_grants":[]})";
		if (!AIChatMcpConfigStore::SaveJsonText(source, &error) ||
			!AIChatMcpConfigStore::SaveJsonText(uiSave, &error, /*preserveMissingMcpServers=*/true)) {
			failures.push_back("seed or WebView-style save failed: " + error);
		}
		else {
			AIChatMcpConfig afterWeb;
			AIChatMcpConfigStore::Load(afterWeb, nullptr);
			const bool preserved = std::any_of(afterWeb.servers.begin(), afterWeb.servers.end(),
				[](const AIChatMcpServerConfig& s) { return s.id == "extern" && s.readOnly; });
			const bool noDup = std::count_if(afterWeb.servers.begin(), afterWeb.servers.end(),
				[](const AIChatMcpServerConfig& s) { return s.id == "extern"; }) == 1;
			if (!preserved) {
				failures.push_back("mcpServers lost after WebView-style (preserve) save");
			}
			if (!noDup) {
				failures.push_back("read-only server duplicated after preserve save");
			}
		}

		// Emulate the native raw-JSON editor dropping mcpServers with preserve=false:
		// the block must actually be deleted.
		if (!AIChatMcpConfigStore::SaveJsonText(source, &error) ||
			!AIChatMcpConfigStore::SaveJsonText(uiSave, &error, /*preserveMissingMcpServers=*/false)) {
			failures.push_back("seed or native-style save failed: " + error);
		}
		else {
			AIChatMcpConfig afterNative;
			AIChatMcpConfigStore::Load(afterNative, nullptr);
			const bool stillThere = std::any_of(afterNative.servers.begin(), afterNative.servers.end(),
				[](const AIChatMcpServerConfig& s) { return s.id == "extern"; });
			if (stillThere) {
				failures.push_back("native-editor deletion of mcpServers was silently reverted");
			}
		}
	}

	// 2. CRLF in a header value is rejected (HTTP header injection).
	{
		const std::string source = R"({"servers":[{"id":"h","name":"H","transport":"streamable_http","url":"http://127.0.0.1:9/mcp","enabled":false,"headers":[{"name":"X-Test","value":"a\r\nAuthorization: bad"}]}]})";
		AIChatMcpConfig parsed;
		std::string error;
		if (AIChatMcpConfigStore::ParseConfigJson(source, parsed, error)) {
			failures.push_back("CRLF header value was accepted");
		}
	}

	// 3. Colliding sanitized ids are made unique instead of overwriting.
	{
		const std::string source = R"({"servers":[
			{"id":"My Server","name":"one","transport":"streamable_http","url":"http://127.0.0.1:9/mcp","enabled":false},
			{"id":"my-server","name":"two","transport":"streamable_http","url":"http://127.0.0.1:9/mcp","enabled":false}
		]})";
		AIChatMcpConfig parsed;
		std::string error;
		if (!AIChatMcpConfigStore::ParseConfigJson(source, parsed, error)) {
			failures.push_back("parse colliding-id config failed: " + error);
		}
		else if (parsed.servers.size() != 2 || parsed.servers[0].id == parsed.servers[1].id) {
			failures.push_back("colliding server ids were not de-duplicated");
		}
	}

	// 4. Explicit streamable_http transport is not overridden by a stray command.
	{
		const std::string source = R"({"mcpServers":{"x":{"transport":"streamable_http","url":"http://127.0.0.1:9/mcp","command":"leftover","enabled":false}}})";
		AIChatMcpConfig parsed;
		std::string error;
		if (!AIChatMcpConfigStore::ParseConfigJson(source, parsed, error)) {
			failures.push_back("parse explicit-transport config failed: " + error);
		}
		else {
			const bool http = std::any_of(parsed.servers.begin(), parsed.servers.end(),
				[](const AIChatMcpServerConfig& s) { return s.id == "x" && s.transport == "streamable_http"; });
			if (!http) {
				failures.push_back("explicit streamable_http overridden to stdio by command");
			}
		}
	}

	outCheck["ok"] = failures.empty();
	if (!failures.empty()) {
		outCheck["error"] = failures.front();
		outCheck["failures"] = failures;
	}
	return failures.empty();
}

bool RunPowerShellRunnerSelfTest(nlohmann::json& outCheck)
{
	outCheck = { {"name", "powershell-runner"}, {"ok", false} };
	const PowerShellRunResult nonZero = PowerShellToolRunner::Run(
		"[Console]::Error.Write('expected-error'); exit 7",
		"",
		10);
	const PowerShellRunResult truncated = PowerShellToolRunner::Run(
		"[Console]::Out.Write('x' * 2200000)",
		"",
		20);
	const PowerShellRunResult timedOut = PowerShellToolRunner::Run(
		"Start-Sleep -Seconds 5",
		"",
		1);
	const PowerShellRunResult childRun = PowerShellToolRunner::Run(
		"$p=Start-Process powershell.exe -WindowStyle Hidden -ArgumentList '-NoLogo','-NoProfile','-NonInteractive','-Command','Start-Sleep -Seconds 30' -PassThru; [Console]::Out.Write($p.Id)",
		"",
		10);

	bool childGone = false;
	unsigned long childProcessId = 0;
	try {
		childProcessId = std::stoul(childRun.stdOut);
	}
	catch (...) {
		childProcessId = 0;
	}
	if (childProcessId != 0) {
		HANDLE child = OpenProcess(SYNCHRONIZE | PROCESS_QUERY_LIMITED_INFORMATION, FALSE, childProcessId);
		if (child == nullptr) {
			childGone = true;
		}
		else {
			childGone = WaitForSingleObject(child, 1000) == WAIT_OBJECT_0;
			CloseHandle(child);
		}
	}

	const bool nonZeroOk = !nonZero.ok && nonZero.exitCode == 7 &&
		nonZero.error.find("code 7") != std::string::npos;
	const bool truncationOk = truncated.ok && truncated.stdOutTruncated &&
		truncated.stdOut.size() <= 2 * 1024 * 1024;
	const bool timeoutOk = !timedOut.ok && timedOut.timedOut;
	const bool childCleanupOk = childRun.ok && childGone;
	outCheck["nonzero_exit"] = {
		{"ok", nonZeroOk},
		{"exit_code", nonZero.exitCode},
		{"error", nonZero.error}
	};
	outCheck["output_truncation"] = {
		{"ok", truncationOk},
		{"stdout_bytes", truncated.stdOut.size()},
		{"stdout_truncated", truncated.stdOutTruncated}
	};
	outCheck["timeout"] = {
		{"ok", timeoutOk},
		{"timed_out", timedOut.timedOut},
		{"exit_code", timedOut.exitCode}
	};
	outCheck["child_cleanup"] = {
		{"ok", childCleanupOk},
		{"child_pid", childProcessId},
		{"child_gone", childGone},
		{"runner_ok", childRun.ok},
		{"runner_error", childRun.error}
	};
	outCheck["ok"] = nonZeroOk && truncationOk && timeoutOk && childCleanupOk;
	return outCheck.value("ok", false);
}

int RunOpenAIIntegrationTestInternal(
	AIProtocolType protocolType,
	const char* apiKey,
	const char* model,
	const char* baseUrl,
	char* buffer,
	int bufferSize)
{
	if (apiKey == nullptr || model == nullptr) {
		return AUTOLINKER_TEST_STRING_INVALID_ARGUMENT;
	}

	const char* defaultBaseUrl = "https://api.openai.com/v1";
	std::string step = "init";
	try {
		AISettings settings = {};
		settings.protocolType = protocolType;
		settings.thinkingLevel = AIThinkingLevel::High;
		settings.baseUrl = (baseUrl != nullptr && baseUrl[0] != '\0') ? baseUrl : defaultBaseUrl;
		settings.apiKey = apiKey;
		settings.model = model;
		settings.timeoutMs = 180000;
		settings.temperature = 0;

		nlohmann::json report = BuildOpenAIIntegrationResultJson(settings);
		report["step"] = step;

		step = "test_connection";
		report["step"] = step;
		const AIResult connectionResult = AIService::TestConnection(settings);
		report["test_connection"] = {
			{"ok", connectionResult.ok},
			{"http_status", connectionResult.httpStatus},
			{"content", connectionResult.content},
			{"error", connectionResult.error}
		};
		if (!connectionResult.ok) {
			report["ok"] = false;
			return CopyStringToBuffer(DumpJsonPrettySafe(report), buffer, bufferSize);
		}

		step = "simple_task";
		report["step"] = step;
		const AIResult simpleTaskResult = AIService::ExecuteTask(
			AITaskKind::TranslateText,
			"只返回这四个字符：测试通过",
			settings);
		report["simple_task"] = {
			{"ok", simpleTaskResult.ok},
			{"http_status", simpleTaskResult.httpStatus},
			{"content", simpleTaskResult.content},
			{"error", simpleTaskResult.error}
		};
		if (!simpleTaskResult.ok) {
			report["ok"] = false;
			return CopyStringToBuffer(DumpJsonPrettySafe(report), buffer, bufferSize);
		}

		step = "tool_chat";
		report["step"] = step;
		std::vector<AIChatMessage> contextMessages;
		contextMessages.push_back({
			"user",
			"你必须先后调用两个工具：先 fetch_url 读取 https://developers.openai.com/api/docs/api-reference/chat/create-chat-completion ，再 extract_web_document 读取 https://developers.openai.com/api/docs/api-reference/responses/create 。完成后仅用一行中文回答，格式必须是：Chat页已读；Responses页已读；工具数=N。",
			"",
			""
		});

		std::vector<std::string> streamedDeltas;
		const AIChatResult toolChatResult = AIService::ExecuteChatWithTools(
			contextMessages,
			settings,
			[](const std::string& toolName, const std::string&, bool& outOk) -> std::string {
				const std::string actualArgs = BuildOpenAIToolArgumentsJson(toolName);
				return ExecuteToolCall(toolName, actualArgs, outOk, false);
			},
			[&streamedDeltas](const std::string& deltaText) {
				if (!deltaText.empty()) {
					streamedDeltas.push_back(deltaText);
				}
			});

		nlohmann::json toolEvents = nlohmann::json::array();
		for (const auto& evt : toolChatResult.toolEvents) {
			toolEvents.push_back({
				{"name", evt.name},
				{"arguments_json", evt.argumentsJson},
				{"result_json", evt.resultJson},
				{"ok", evt.ok}
			});
		}
		report["tool_chat"] = {
			{"ok", toolChatResult.ok},
			{"cancelled", toolChatResult.cancelled},
			{"http_status", toolChatResult.httpStatus},
			{"content", toolChatResult.content},
			{"reasoning_content_present", !toolChatResult.reasoningContent.empty()},
			{"reasoning_content_size", toolChatResult.reasoningContent.size()},
			{"error", toolChatResult.error},
			{"tool_events", std::move(toolEvents)},
			{"stream_chunk_count", streamedDeltas.size()},
			{"hidden_context_message_count", toolChatResult.contextPrefixRawMessagesUtf8.size()}
		};
		if (!toolChatResult.ok) {
			report["ok"] = false;
			return CopyStringToBuffer(DumpJsonPrettySafe(report), buffer, bufferSize);
		}

		step = "followup_chat";
		report["step"] = step;
		std::vector<AIChatMessage> followupMessages = BuildFollowupMessagesFromChatResult(contextMessages, toolChatResult);
		followupMessages.push_back({
			"assistant",
			toolChatResult.content,
			toolChatResult.reasoningContent,
			""
		});
		followupMessages.push_back({
			"user",
			"只回答：上一轮你实际调用了几个工具？输出阿拉伯数字。",
			"",
			""
		});

		const AIChatResult followupResult = AIService::ExecuteChatWithTools(
			followupMessages,
			settings,
			[](const std::string& toolName, const std::string&, bool& outOk) -> std::string {
				const std::string actualArgs = BuildOpenAIToolArgumentsJson(toolName);
				return ExecuteToolCall(toolName, actualArgs, outOk, false);
			});
		report["followup_chat"] = {
			{"ok", followupResult.ok},
			{"cancelled", followupResult.cancelled},
			{"http_status", followupResult.httpStatus},
			{"content", followupResult.content},
			{"reasoning_content_present", !followupResult.reasoningContent.empty()},
			{"reasoning_content_size", followupResult.reasoningContent.size()},
			{"error", followupResult.error},
			{"tool_event_count", followupResult.toolEvents.size()}
		};

		report["ok"] =
			connectionResult.ok &&
			simpleTaskResult.ok &&
			toolChatResult.ok &&
			followupResult.ok &&
			toolChatResult.toolEvents.size() >= 2;
		return CopyStringToBuffer(DumpJsonPrettySafe(report), buffer, bufferSize);
	}
	catch (const std::exception& ex) {
		nlohmann::json report = {
			{"ok", false},
			{"provider", "openai"},
			{"model", model},
			{"base_url", (baseUrl != nullptr && baseUrl[0] != '\0') ? baseUrl : defaultBaseUrl},
			{"protocol", protocolType == AIProtocolType::OpenAIResponses ? "openai_responses" : "openai"},
			{"step", step},
			{"error", std::string("exception: ") + ex.what()}
		};
		return CopyStringToBuffer(DumpJsonPrettySafe(report), buffer, bufferSize);
	}
	catch (...) {
		nlohmann::json report = {
			{"ok", false},
			{"provider", "openai"},
			{"model", model},
			{"base_url", (baseUrl != nullptr && baseUrl[0] != '\0') ? baseUrl : defaultBaseUrl},
			{"protocol", protocolType == AIProtocolType::OpenAIResponses ? "openai_responses" : "openai"},
			{"step", step},
			{"error", "unknown exception"}
		};
		return CopyStringToBuffer(DumpJsonPrettySafe(report), buffer, bufferSize);
	}
}

int RunGeminiIntegrationTestInternal(
	const char* apiKey,
	const char* model,
	const char* baseUrl,
	char* buffer,
	int bufferSize)
{
	if (apiKey == nullptr || model == nullptr) {
		return AUTOLINKER_TEST_STRING_INVALID_ARGUMENT;
	}

	const char* defaultBaseUrl = "https://generativelanguage.googleapis.com";
	std::string step = "init";
	try {
		AISettings settings = {};
		settings.protocolType = AIProtocolType::Gemini;
		settings.thinkingLevel = AIThinkingLevel::Off;
		settings.baseUrl = (baseUrl != nullptr && baseUrl[0] != '\0') ? baseUrl : defaultBaseUrl;
		settings.apiKey = apiKey;
		settings.model = model;
		settings.timeoutMs = 180000;
		settings.temperature = 0;

		nlohmann::json report = BuildGeminiIntegrationResultJson(settings);
		report["step"] = step;

		step = "test_connection";
		report["step"] = step;
		const AIResult connectionResult = AIService::TestConnection(settings);
		report["test_connection"] = {
			{"ok", connectionResult.ok},
			{"http_status", connectionResult.httpStatus},
			{"content", connectionResult.content},
			{"error", connectionResult.error}
		};
		if (!connectionResult.ok) {
			report["ok"] = false;
			return CopyStringToBuffer(DumpJsonPrettySafe(report), buffer, bufferSize);
		}

		step = "simple_task";
		report["step"] = step;
		const AIResult simpleTaskResult = AIService::ExecuteTask(
			AITaskKind::TranslateText,
			"只返回这四个字符：测试通过",
			settings);
		report["simple_task"] = {
			{"ok", simpleTaskResult.ok},
			{"http_status", simpleTaskResult.httpStatus},
			{"content", simpleTaskResult.content},
			{"error", simpleTaskResult.error}
		};
		if (!simpleTaskResult.ok) {
			report["ok"] = false;
			return CopyStringToBuffer(DumpJsonPrettySafe(report), buffer, bufferSize);
		}

		step = "tool_chat";
		report["step"] = step;
		std::vector<AIChatMessage> contextMessages;
		contextMessages.push_back({
			"user",
			"你必须先后调用两个工具：先 fetch_url 读取 https://ai.google.dev/gemini-api/docs/models/gemini ，再 extract_web_document 读取 https://ai.google.dev/gemini-api/docs/function-calling 。完成后仅用一行中文回答，格式必须是：模型页已读；函数调用页已读；工具数=N。",
			"",
			""
		});

		std::vector<std::string> streamedDeltas;
		const AIChatResult toolChatResult = AIService::ExecuteChatWithTools(
			contextMessages,
			settings,
			[](const std::string& toolName, const std::string&, bool& outOk) -> std::string {
				const std::string actualArgs = BuildGeminiToolArgumentsJson(toolName);
				return ExecuteToolCall(toolName, actualArgs, outOk, false);
			},
			[&streamedDeltas](const std::string& deltaText) {
				if (!deltaText.empty()) {
					streamedDeltas.push_back(deltaText);
				}
			});

		nlohmann::json toolEvents = nlohmann::json::array();
		bool allToolEventsOk = true;
		for (const auto& evt : toolChatResult.toolEvents) {
			allToolEventsOk = allToolEventsOk && evt.ok;
			toolEvents.push_back({
				{"name", evt.name},
				{"arguments_json", evt.argumentsJson},
				{"result_json", evt.resultJson},
				{"ok", evt.ok}
			});
		}
		report["tool_chat"] = {
			{"ok", toolChatResult.ok},
			{"cancelled", toolChatResult.cancelled},
			{"http_status", toolChatResult.httpStatus},
			{"content", toolChatResult.content},
			{"reasoning_content_present", !toolChatResult.reasoningContent.empty()},
			{"reasoning_content_size", toolChatResult.reasoningContent.size()},
			{"error", toolChatResult.error},
			{"tool_events", std::move(toolEvents)},
			{"all_tool_events_ok", allToolEventsOk},
			{"stream_chunk_count", streamedDeltas.size()},
			{"hidden_context_message_count", toolChatResult.contextPrefixRawMessagesUtf8.size()}
		};
		if (!toolChatResult.ok) {
			report["ok"] = false;
			return CopyStringToBuffer(DumpJsonPrettySafe(report), buffer, bufferSize);
		}

		step = "followup_chat";
		report["step"] = step;
		std::vector<AIChatMessage> followupMessages = BuildFollowupMessagesFromChatResult(contextMessages, toolChatResult);
		followupMessages.push_back({
			"assistant",
			toolChatResult.content,
			toolChatResult.reasoningContent,
			""
		});
		followupMessages.push_back({
			"user",
			"只回答：上一轮你实际调用了几个工具？输出阿拉伯数字。",
			"",
			""
		});

		const AIChatResult followupResult = AIService::ExecuteChatWithTools(
			followupMessages,
			settings,
			[](const std::string& toolName, const std::string&, bool& outOk) -> std::string {
				const std::string actualArgs = BuildGeminiToolArgumentsJson(toolName);
				return ExecuteToolCall(toolName, actualArgs, outOk, false);
			});
		report["followup_chat"] = {
			{"ok", followupResult.ok},
			{"cancelled", followupResult.cancelled},
			{"http_status", followupResult.httpStatus},
			{"content", followupResult.content},
			{"reasoning_content_present", !followupResult.reasoningContent.empty()},
			{"reasoning_content_size", followupResult.reasoningContent.size()},
			{"error", followupResult.error},
			{"tool_event_count", followupResult.toolEvents.size()}
		};

		report["ok"] =
			connectionResult.ok &&
			simpleTaskResult.ok &&
			toolChatResult.ok &&
			followupResult.ok &&
			toolChatResult.toolEvents.size() >= 2 &&
			allToolEventsOk;
		return CopyStringToBuffer(DumpJsonPrettySafe(report), buffer, bufferSize);
	}
	catch (const std::exception& ex) {
		nlohmann::json report = {
			{"ok", false},
			{"provider", "gemini"},
			{"model", model},
			{"base_url", (baseUrl != nullptr && baseUrl[0] != '\0') ? baseUrl : defaultBaseUrl},
			{"protocol", "gemini"},
			{"step", step},
			{"error", std::string("exception: ") + ex.what()}
		};
		return CopyStringToBuffer(DumpJsonPrettySafe(report), buffer, bufferSize);
	}
	catch (...) {
		nlohmann::json report = {
			{"ok", false},
			{"provider", "gemini"},
			{"model", model},
			{"base_url", (baseUrl != nullptr && baseUrl[0] != '\0') ? baseUrl : defaultBaseUrl},
			{"protocol", "gemini"},
			{"step", step},
			{"error", "unknown exception"}
		};
		return CopyStringToBuffer(DumpJsonPrettySafe(report), buffer, bufferSize);
	}
}

int RunClaudeIntegrationTestInternal(
	const char* apiKey,
	const char* model,
	const char* baseUrl,
	char* buffer,
	int bufferSize)
{
	if (apiKey == nullptr || model == nullptr) {
		return AUTOLINKER_TEST_STRING_INVALID_ARGUMENT;
	}

	const char* defaultBaseUrl = "https://api.anthropic.com";
	std::string step = "init";
	try {
		AISettings settings = {};
		settings.protocolType = AIProtocolType::Claude;
		settings.thinkingLevel = AIThinkingLevel::Off;
		settings.baseUrl = (baseUrl != nullptr && baseUrl[0] != '\0') ? baseUrl : defaultBaseUrl;
		settings.apiKey = apiKey;
		settings.model = model;
		settings.timeoutMs = 180000;
		settings.temperature = 0;

		nlohmann::json report = BuildClaudeIntegrationResultJson(settings);
		report["step"] = step;

		step = "test_connection";
		report["step"] = step;
		const AIResult connectionResult = AIService::TestConnection(settings);
		report["test_connection"] = {
			{"ok", connectionResult.ok},
			{"http_status", connectionResult.httpStatus},
			{"content", connectionResult.content},
			{"error", connectionResult.error}
		};
		if (!connectionResult.ok) {
			report["ok"] = false;
			return CopyStringToBuffer(DumpJsonPrettySafe(report), buffer, bufferSize);
		}

		step = "simple_task";
		report["step"] = step;
		const AIResult simpleTaskResult = AIService::ExecuteTask(
			AITaskKind::TranslateText,
			"只返回这四个字符：测试通过",
			settings);
		report["simple_task"] = {
			{"ok", simpleTaskResult.ok},
			{"http_status", simpleTaskResult.httpStatus},
			{"content", simpleTaskResult.content},
			{"error", simpleTaskResult.error}
		};
		if (!simpleTaskResult.ok) {
			report["ok"] = false;
			return CopyStringToBuffer(DumpJsonPrettySafe(report), buffer, bufferSize);
		}

		step = "tool_chat";
		report["step"] = step;
		std::vector<AIChatMessage> contextMessages;
		contextMessages.push_back({
			"user",
			"你必须先后调用两个工具：先 fetch_url 读取 https://docs.anthropic.com/en/docs/about-claude/models/overview ，再 extract_web_document 读取 https://docs.anthropic.com/en/docs/agents-and-tools/tool-use/overview 。完成后仅用一行中文回答，格式必须是：模型页已读；工具页已读；工具数=N。",
			"",
			""
		});

		std::vector<std::string> streamedDeltas;
		const AIChatResult toolChatResult = AIService::ExecuteChatWithTools(
			contextMessages,
			settings,
			[](const std::string& toolName, const std::string&, bool& outOk) -> std::string {
				const std::string actualArgs = BuildClaudeToolArgumentsJson(toolName);
				return ExecuteToolCall(toolName, actualArgs, outOk, false);
			},
			[&streamedDeltas](const std::string& deltaText) {
				if (!deltaText.empty()) {
					streamedDeltas.push_back(deltaText);
				}
			});

		nlohmann::json toolEvents = nlohmann::json::array();
		bool allToolEventsOk = true;
		for (const auto& evt : toolChatResult.toolEvents) {
			allToolEventsOk = allToolEventsOk && evt.ok;
			toolEvents.push_back({
				{"name", evt.name},
				{"arguments_json", evt.argumentsJson},
				{"result_json", evt.resultJson},
				{"ok", evt.ok}
			});
		}
		const bool hiddenContextOk = toolChatResult.contextPrefixRawMessagesUtf8.size() >= 2;
		report["tool_chat"] = {
			{"ok", toolChatResult.ok},
			{"cancelled", toolChatResult.cancelled},
			{"http_status", toolChatResult.httpStatus},
			{"content", toolChatResult.content},
			{"reasoning_content_present", !toolChatResult.reasoningContent.empty()},
			{"reasoning_content_size", toolChatResult.reasoningContent.size()},
			{"error", toolChatResult.error},
			{"tool_events", std::move(toolEvents)},
			{"all_tool_events_ok", allToolEventsOk},
			{"hidden_context_ok", hiddenContextOk},
			{"stream_chunk_count", streamedDeltas.size()},
			{"hidden_context_message_count", toolChatResult.contextPrefixRawMessagesUtf8.size()}
		};
		if (!toolChatResult.ok) {
			report["ok"] = false;
			return CopyStringToBuffer(DumpJsonPrettySafe(report), buffer, bufferSize);
		}

		step = "followup_chat";
		report["step"] = step;
		std::vector<AIChatMessage> followupMessages = BuildFollowupMessagesFromChatResult(contextMessages, toolChatResult);
		followupMessages.push_back({
			"assistant",
			toolChatResult.content,
			toolChatResult.reasoningContent,
			""
		});
		followupMessages.push_back({
			"user",
			"只回答：上一轮你实际调用了几个工具？输出阿拉伯数字。",
			"",
			""
		});

		const AIChatResult followupResult = AIService::ExecuteChatWithTools(
			followupMessages,
			settings,
			[](const std::string& toolName, const std::string&, bool& outOk) -> std::string {
				const std::string actualArgs = BuildClaudeToolArgumentsJson(toolName);
				return ExecuteToolCall(toolName, actualArgs, outOk, false);
			});
		report["followup_chat"] = {
			{"ok", followupResult.ok},
			{"cancelled", followupResult.cancelled},
			{"http_status", followupResult.httpStatus},
			{"content", followupResult.content},
			{"reasoning_content_present", !followupResult.reasoningContent.empty()},
			{"reasoning_content_size", followupResult.reasoningContent.size()},
			{"error", followupResult.error},
			{"tool_event_count", followupResult.toolEvents.size()}
		};

		report["ok"] =
			connectionResult.ok &&
			simpleTaskResult.ok &&
			toolChatResult.ok &&
			followupResult.ok &&
			toolChatResult.toolEvents.size() >= 2 &&
			allToolEventsOk &&
			hiddenContextOk;
		return CopyStringToBuffer(DumpJsonPrettySafe(report), buffer, bufferSize);
	}
	catch (const std::exception& ex) {
		nlohmann::json report = {
			{"ok", false},
			{"provider", "claude"},
			{"model", model},
			{"base_url", (baseUrl != nullptr && baseUrl[0] != '\0') ? baseUrl : defaultBaseUrl},
			{"protocol", "claude"},
			{"step", step},
			{"error", std::string("exception: ") + ex.what()}
		};
		return CopyStringToBuffer(DumpJsonPrettySafe(report), buffer, bufferSize);
	}
	catch (...) {
		nlohmann::json report = {
			{"ok", false},
			{"provider", "claude"},
			{"model", model},
			{"base_url", (baseUrl != nullptr && baseUrl[0] != '\0') ? baseUrl : defaultBaseUrl},
			{"protocol", "claude"},
			{"step", step},
			{"error", "unknown exception"}
		};
		return CopyStringToBuffer(DumpJsonPrettySafe(report), buffer, bufferSize);
	}
}

}

extern "C" bool AutoLinkerTest_CompareVersion(const char* left, const char* right, int* outResult)
{
	if (left == nullptr || right == nullptr || outResult == nullptr) {
		return false;
	}

	const Version leftVersion(left);
	const Version rightVersion(right);
	if (leftVersion < rightVersion) {
		*outResult = -1;
	}
	else if (leftVersion > rightVersion) {
		*outResult = 1;
	}
	else {
		*outResult = 0;
	}

	return true;
}

extern "C" int AutoLinkerTest_GetLinkerOutFileName(const char* commandLine, char* buffer, int bufferSize)
{
	if (commandLine == nullptr) {
		return AUTOLINKER_TEST_STRING_INVALID_ARGUMENT;
	}

	return CopyStringToBuffer(GetLinkerCommandOutFileName(commandLine), buffer, bufferSize);
}

extern "C" int AutoLinkerTest_GetLinkerKrnlnFileName(const char* commandLine, char* buffer, int bufferSize)
{
	if (commandLine == nullptr) {
		return AUTOLINKER_TEST_STRING_INVALID_ARGUMENT;
	}

	return CopyStringToBuffer(GetLinkerCommandKrnlnFileName(commandLine), buffer, bufferSize);
}

extern "C" int AutoLinkerTest_ExtractBetweenDashes(const char* text, char* buffer, int bufferSize)
{
	if (text == nullptr) {
		return AUTOLINKER_TEST_STRING_INVALID_ARGUMENT;
	}

	return CopyStringToBuffer(ExtractBetweenDashes(text), buffer, bufferSize);
}

extern "C" int AutoLinkerTest_GetVersionText(char* buffer, int bufferSize)
{
	return CopyStringToBuffer(AUTOLINKER_VERSION, buffer, bufferSize);
}

extern "C" int AutoLinkerTest_RunGameAnalyticsSelfTest(char* buffer, int bufferSize)
{
	return CopyStringToBuffer(GameAnalyticsClient::BuildSelfTestReportJson(), buffer, bufferSize);
}

extern "C" int AutoLinkerTest_RunAIChatMcpSelfTest(char* buffer, int bufferSize)
{
	nlohmann::json report = nlohmann::json::parse(AIChatMcpClient::BuildSelfTestReportJson(), nullptr, false);
	if (report.is_discarded() || !report.is_object()) {
		report = {
			{"ok", false},
			{"name", "mcp-self-test"},
			{"checks", nlohmann::json::array()}
		};
	}
	if (!report.contains("checks") || !report["checks"].is_array()) {
		report["checks"] = nlohmann::json::array();
	}

	nlohmann::json publicCatalogCheck = {
		{"name", "public_catalog_excludes_mcp"},
		{"ok", true}
	};
	try {
		const nlohmann::json catalog = nlohmann::json::parse(AIService::BuildPublicToolCatalogJson(), nullptr, false);
		if (catalog.is_discarded() || !catalog.is_array()) {
			publicCatalogCheck["ok"] = false;
			publicCatalogCheck["error"] = "public catalog is not array";
		}
		else {
			bool hasReadCodeItem = false;
			bool hasRequiredRefreshDescription = false;
			bool editCasSchemaOk = false;
			bool multiEditCasSchemaOk = false;
			bool diffCasSchemaOk = false;
			bool restoreCasSchemaOk = false;
			bool hiddenDependencyTools = true;
			for (const auto& item : catalog) {
				if (item.is_object() && AIChatMcpClient::IsMcpModelToolName(item.value("name", std::string()))) {
					publicCatalogCheck["ok"] = false;
					publicCatalogCheck["error"] = "public catalog contains MCP tool";
					break;
				}
				if (!item.is_object()) {
					continue;
				}
				const std::string name = item.value("name", std::string());
				if (name == "read_code_item") {
					hasReadCodeItem = true;
				}
				const nlohmann::json properties = item.value("inputSchema", nlohmann::json::object())
					.value("properties", nlohmann::json::object());
				editCasSchemaOk = editCasSchemaOk || (name == "edit_file" && properties.contains("expected_base_hash"));
				multiEditCasSchemaOk = multiEditCasSchemaOk || (name == "multi_edit_file" && properties.contains("expected_base_hash"));
				diffCasSchemaOk = diffCasSchemaOk || (name == "diff_file" && properties.contains("expected_base_hash"));
				restoreCasSchemaOk = restoreCasSchemaOk || (name == "restore_file_snapshot" && properties.contains("expected_current_hash"));
				if (name == "refresh_dependency_catalog" ||
					name == "search_available_modules" ||
					name == "add_module_to_project") {
					hiddenDependencyTools = false;
				}
				if (name == "refresh_workspace_mirror" &&
					item.value("description", std::string()).find("MUST") != std::string::npos) {
					hasRequiredRefreshDescription = true;
				}
			}
			const bool casSchemasOk = editCasSchemaOk && multiEditCasSchemaOk && diffCasSchemaOk && restoreCasSchemaOk;
			if (!hasReadCodeItem || !hasRequiredRefreshDescription || !casSchemasOk || !hiddenDependencyTools) {
				publicCatalogCheck["ok"] = false;
				publicCatalogCheck["error"] = "public catalog visibility or CAS schema invariant failed";
			}
			publicCatalogCheck["has_read_code_item"] = hasReadCodeItem;
			publicCatalogCheck["has_required_refresh_description"] = hasRequiredRefreshDescription;
			publicCatalogCheck["cas_schemas_ok"] = casSchemasOk;
			publicCatalogCheck["dependency_tools_hidden"] = hiddenDependencyTools;
		}
	}
	catch (const std::exception& ex) {
		publicCatalogCheck["ok"] = false;
		publicCatalogCheck["error"] = ex.what();
	}
	report["checks"].push_back(publicCatalogCheck);

	nlohmann::json toolPolicyCheck;
	RunAIChatToolPolicySelfTest(toolPolicyCheck);
	report["checks"].push_back(std::move(toolPolicyCheck));

	nlohmann::json optimizationCheck = nlohmann::json::parse(
		AIService::BuildAgentOptimizationSelfTestJson(),
		nullptr,
		false);
	if (optimizationCheck.is_discarded() || !optimizationCheck.is_object()) {
		optimizationCheck = {
			{"name", "agent-optimization-self-test"},
			{"ok", false},
			{"error", "invalid self-test json"}
		};
	}
	report["checks"].push_back(std::move(optimizationCheck));

	nlohmann::json paginationCheck = nlohmann::json::parse(
		WorkspaceFileTools::BuildPaginationSelfTestJson(),
		nullptr,
		false);
	if (paginationCheck.is_discarded() || !paginationCheck.is_object()) {
		paginationCheck = {
			{"name", "workspace-file-pagination"},
			{"ok", false},
			{"error", "invalid self-test json"}
		};
	}
	report["checks"].push_back(std::move(paginationCheck));

	nlohmann::json mirrorSafetyCheck = nlohmann::json::parse(
		WorkspaceMirror::BuildSelfTestReportJson(),
		nullptr,
		false);
	if (mirrorSafetyCheck.is_discarded() || !mirrorSafetyCheck.is_object()) {
		mirrorSafetyCheck = {
			{"name", "workspace-mirror-safety"},
			{"ok", false},
			{"error", "invalid self-test json"}
		};
	}
	report["checks"].push_back(std::move(mirrorSafetyCheck));

	nlohmann::json compileFingerprintCheck = nlohmann::json::parse(
		BuildCompileArtifactFingerprintSelfTestJson(),
		nullptr,
		false);
	if (compileFingerprintCheck.is_discarded() || !compileFingerprintCheck.is_object()) {
		compileFingerprintCheck = {
			{"name", "compile-artifact-fingerprint"},
			{"ok", false},
			{"error", "invalid self-test json"}
		};
	}
	report["checks"].push_back(std::move(compileFingerprintCheck));

	const std::string normalizedHashA = BuildStableTextHashForRealCode("a\r\nb");
	const std::string normalizedHashB = BuildStableTextHashForRealCode("a\nb");
	const std::string differentHash = BuildStableTextHashForRealCode("a\nb\nchanged");
	report["checks"].push_back({
		{"name", "real-page-sha256-hash"},
		{"ok", normalizedHashA.size() == 64 && normalizedHashA == normalizedHashB && normalizedHashA != differentHash},
		{"hash_length", normalizedHashA.size()},
		{"line_break_normalized", normalizedHashA == normalizedHashB},
		{"different_text_differs", normalizedHashA != differentHash}
	});

	nlohmann::json refreshGateCheck = nlohmann::json::parse(
		LocalMcpServer::BuildWorkspaceRefreshGateSelfTestJson(),
		nullptr,
		false);
	if (refreshGateCheck.is_discarded() || !refreshGateCheck.is_object()) {
		refreshGateCheck = {
			{"name", "external_mcp_workspace_refresh_gate"},
			{"ok", false},
			{"error", "invalid self-test json"}
		};
	}
	report["checks"].push_back(std::move(refreshGateCheck));

	nlohmann::json mockRoundtripCheck;
	RunMcpMockRoundtripSelfTest(mockRoundtripCheck);
	report["checks"].push_back(mockRoundtripCheck);

	nlohmann::json mockStdioRoundtripCheck;
	RunMcpStdioRoundtripSelfTest(mockStdioRoundtripCheck);
	report["checks"].push_back(mockStdioRoundtripCheck);

	nlohmann::json powerShellRunnerCheck;
	RunPowerShellRunnerSelfTest(powerShellRunnerCheck);
	report["checks"].push_back(powerShellRunnerCheck);

	for (const std::string& webSelfTest : {
			WebDocumentClient::BuildSelfTestReportJson(),
			WebDocumentExtractor::BuildSelfTestReportJson() }) {
		nlohmann::json webCheck = nlohmann::json::parse(webSelfTest, nullptr, false);
		if (webCheck.is_discarded() || !webCheck.is_object()) {
			webCheck = {
				{"name", "web-document-self-test"},
				{"ok", false},
				{"error", "invalid self-test json"}
			};
		}
		report["checks"].push_back(std::move(webCheck));
	}

	nlohmann::json configInvariantsCheck;
	RunMcpConfigInvariantsSelfTest(configInvariantsCheck);
	report["checks"].push_back(configInvariantsCheck);

	bool ok = report.value("ok", false);
	for (const auto& check : report["checks"]) {
		ok = ok && check.value("ok", false);
	}
	report["ok"] = ok;
	return CopyStringToBuffer(DumpJsonPrettySafe(report), buffer, bufferSize);
}

extern "C" int AutoLinkerTest_RunDeepSeekModelIntegrationTest(
	const char* apiKey,
	const char* model,
	const char* baseUrl,
	char* buffer,
	int bufferSize)
{
	if (apiKey == nullptr || model == nullptr) {
		return AUTOLINKER_TEST_STRING_INVALID_ARGUMENT;
	}

	std::string step = "init";
	try {
		AISettings settings = {};
		settings.protocolType = AIProtocolType::OpenAI;
		settings.thinkingLevel = AIThinkingLevel::High;
		settings.baseUrl = (baseUrl != nullptr && baseUrl[0] != '\0') ? baseUrl : "https://api.deepseek.com";
		settings.apiKey = apiKey;
		settings.model = model;
		settings.timeoutMs = 180000;
		settings.temperature = 0;

		nlohmann::json report = BuildDeepSeekIntegrationResultJson(settings);
		report["step"] = step;

		step = "test_connection";
		report["step"] = step;
		const AIResult connectionResult = AIService::TestConnection(settings);
		report["test_connection"] = {
			{"ok", connectionResult.ok},
			{"http_status", connectionResult.httpStatus},
			{"content", connectionResult.content},
			{"error", connectionResult.error}
		};
		if (!connectionResult.ok) {
			report["ok"] = false;
			return CopyStringToBuffer(DumpJsonPrettySafe(report), buffer, bufferSize);
		}

		step = "simple_task";
		report["step"] = step;
		const AIResult simpleTaskResult = AIService::ExecuteTask(
			AITaskKind::TranslateText,
			"只返回这四个字符：测试通过",
			settings);
		report["simple_task"] = {
			{"ok", simpleTaskResult.ok},
			{"http_status", simpleTaskResult.httpStatus},
			{"content", simpleTaskResult.content},
			{"error", simpleTaskResult.error}
		};
		if (!simpleTaskResult.ok) {
			report["ok"] = false;
			return CopyStringToBuffer(DumpJsonPrettySafe(report), buffer, bufferSize);
		}

		step = "tool_chat";
		report["step"] = step;
		std::vector<AIChatMessage> contextMessages;
		contextMessages.push_back({
			"user",
			"你必须先后调用两个工具：先 fetch_url 读取 https://api-docs.deepseek.com/quick_start/rate_limit ，再 extract_web_document 读取 https://api-docs.deepseek.com/guides/thinking_mode 。完成后仅用一行中文回答，格式必须是：限速页已读；思考页已读；工具数=N。",
			"",
			""
		});

		std::vector<std::string> streamedDeltas;
		const AIChatResult toolChatResult = AIService::ExecuteChatWithTools(
			contextMessages,
			settings,
			[](const std::string& toolName, const std::string&, bool& outOk) -> std::string {
				const std::string actualArgs = BuildDeepSeekToolArgumentsJson(toolName);
				return ExecuteToolCall(toolName, actualArgs, outOk, false);
			},
			[&streamedDeltas](const std::string& deltaText) {
				if (!deltaText.empty()) {
					streamedDeltas.push_back(deltaText);
				}
			});

		nlohmann::json toolEvents = nlohmann::json::array();
		for (const auto& evt : toolChatResult.toolEvents) {
			toolEvents.push_back({
				{"name", evt.name},
				{"arguments_json", evt.argumentsJson},
				{"result_json", evt.resultJson},
				{"ok", evt.ok}
			});
		}
		report["tool_chat"] = {
			{"ok", toolChatResult.ok},
			{"cancelled", toolChatResult.cancelled},
			{"http_status", toolChatResult.httpStatus},
			{"content", toolChatResult.content},
			{"reasoning_content_present", !toolChatResult.reasoningContent.empty()},
			{"reasoning_content_size", toolChatResult.reasoningContent.size()},
			{"error", toolChatResult.error},
			{"tool_events", std::move(toolEvents)},
			{"stream_chunk_count", streamedDeltas.size()},
			{"hidden_context_message_count", toolChatResult.contextPrefixRawMessagesUtf8.size()}
		};
		if (!toolChatResult.ok) {
			report["ok"] = false;
			return CopyStringToBuffer(DumpJsonPrettySafe(report), buffer, bufferSize);
		}

		step = "followup_chat";
		report["step"] = step;
		std::vector<AIChatMessage> followupMessages;
		followupMessages.push_back(contextMessages.front());
		for (const auto& rawMessageJsonUtf8 : toolChatResult.contextPrefixRawMessagesUtf8) {
			nlohmann::json parsed;
			try {
				parsed = nlohmann::json::parse(rawMessageJsonUtf8);
			}
			catch (...) {
				continue;
			}
			if (!parsed.is_object()) {
				continue;
			}

			const std::string role = parsed.value("role", std::string());
			if (role == "assistant") {
				followupMessages.push_back({
					"assistant",
					ExtractMessageContentText(parsed),
					parsed.value("reasoning_content", std::string()),
					rawMessageJsonUtf8
				});
			}
			else if (role == "tool") {
				followupMessages.push_back({
					"tool",
					ExtractMessageContentText(parsed),
					"",
					rawMessageJsonUtf8
				});
			}
		}
		followupMessages.push_back({
			"assistant",
			toolChatResult.content,
			toolChatResult.reasoningContent,
			""
		});
		followupMessages.push_back({
			"user",
			"只回答：上一轮你实际调用了几个工具？输出阿拉伯数字。",
			"",
			""
		});

		const AIChatResult followupResult = AIService::ExecuteChatWithTools(
			followupMessages,
			settings,
			[](const std::string& toolName, const std::string&, bool& outOk) -> std::string {
				const std::string actualArgs = BuildDeepSeekToolArgumentsJson(toolName);
				return ExecuteToolCall(toolName, actualArgs, outOk, false);
			});
		report["followup_chat"] = {
			{"ok", followupResult.ok},
			{"cancelled", followupResult.cancelled},
			{"http_status", followupResult.httpStatus},
			{"content", followupResult.content},
			{"reasoning_content_present", !followupResult.reasoningContent.empty()},
			{"reasoning_content_size", followupResult.reasoningContent.size()},
			{"error", followupResult.error},
			{"tool_event_count", followupResult.toolEvents.size()}
		};

		report["ok"] =
			connectionResult.ok &&
			simpleTaskResult.ok &&
			toolChatResult.ok &&
			followupResult.ok &&
			toolChatResult.toolEvents.size() >= 2;
		return CopyStringToBuffer(DumpJsonPrettySafe(report), buffer, bufferSize);
	}
	catch (const std::exception& ex) {
		nlohmann::json report = {
			{"ok", false},
			{"provider", "deepseek"},
			{"model", model},
			{"base_url", (baseUrl != nullptr && baseUrl[0] != '\0') ? baseUrl : "https://api.deepseek.com"},
			{"step", step},
			{"error", std::string("exception: ") + ex.what()}
		};
		return CopyStringToBuffer(DumpJsonPrettySafe(report), buffer, bufferSize);
	}
	catch (...) {
		nlohmann::json report = {
			{"ok", false},
			{"provider", "deepseek"},
			{"model", model},
			{"base_url", (baseUrl != nullptr && baseUrl[0] != '\0') ? baseUrl : "https://api.deepseek.com"},
			{"step", step},
			{"error", "unknown exception"}
		};
		return CopyStringToBuffer(DumpJsonPrettySafe(report), buffer, bufferSize);
	}
}

extern "C" int AutoLinkerTest_RunDeepSeekConnectionOnly(
	const char* apiKey,
	const char* model,
	const char* baseUrl,
	char* buffer,
	int bufferSize)
{
	if (apiKey == nullptr || model == nullptr) {
		return AUTOLINKER_TEST_STRING_INVALID_ARGUMENT;
	}

	AISettings settings = {};
	settings.protocolType = AIProtocolType::OpenAI;
	settings.thinkingLevel = AIThinkingLevel::High;
	settings.baseUrl = (baseUrl != nullptr && baseUrl[0] != '\0') ? baseUrl : "https://api.deepseek.com";
	settings.apiKey = apiKey;
	settings.model = model;
	settings.timeoutMs = 180000;
	settings.temperature = 0;

	const AIResult connectionResult = AIService::TestConnection(settings);
	nlohmann::json report = BuildDeepSeekIntegrationResultJson(settings);
	report["ok"] = connectionResult.ok;
	report["test_connection"] = {
		{"ok", connectionResult.ok},
		{"http_status", connectionResult.httpStatus},
		{"content", connectionResult.content},
		{"error", connectionResult.error}
	};
	return CopyStringToBuffer(DumpJsonPrettySafe(report), buffer, bufferSize);
}

extern "C" int AutoLinkerTest_RunOpenAIChatIntegrationTest(
	const char* apiKey,
	const char* model,
	const char* baseUrl,
	char* buffer,
	int bufferSize)
{
	return RunOpenAIIntegrationTestInternal(
		AIProtocolType::OpenAI,
		apiKey,
		model,
		baseUrl,
		buffer,
		bufferSize);
}

extern "C" int AutoLinkerTest_RunOpenAIResponsesIntegrationTest(
	const char* apiKey,
	const char* model,
	const char* baseUrl,
	char* buffer,
	int bufferSize)
{
	return RunOpenAIIntegrationTestInternal(
		AIProtocolType::OpenAIResponses,
		apiKey,
		model,
		baseUrl,
		buffer,
		bufferSize);
}

extern "C" int AutoLinkerTest_RunGeminiIntegrationTest(
	const char* apiKey,
	const char* model,
	const char* baseUrl,
	char* buffer,
	int bufferSize)
{
	return RunGeminiIntegrationTestInternal(
		apiKey,
		model,
		baseUrl,
		buffer,
		bufferSize);
}

extern "C" int AutoLinkerTest_RunClaudeIntegrationTest(
	const char* apiKey,
	const char* model,
	const char* baseUrl,
	char* buffer,
	int bufferSize)
{
	return RunClaudeIntegrationTestInternal(
		apiKey,
		model,
		baseUrl,
		buffer,
		bufferSize);
}
