#include "LocalMcpServer.h"

#include <winsock2.h>
#include <ws2tcpip.h>

#include <algorithm>
#include <atomic>
#include <cctype>
#include <chrono>
#include <condition_variable>
#include <cstdint>
#include <cstdlib>
#include <deque>
#include <filesystem>
#include <format>
#include <mutex>
#include <string>
#include <thread>
#include <unordered_map>
#include <vector>

#include "..\\thirdparty\\json.hpp"

#include "AIChatFeature.h"
#include "AIChatToolRegistry.h"
#include "AIChatTooling.h"
#include "AIService.h"
#include "Global.h"
#include "IDEFacade.h"
#include "LocalMcpInstanceRegistry.h"
#include "Logger.h"
#include "PathHelper.h"
#include "WorkspaceMirror.h"

#pragma comment(lib, "Ws2_32.lib")

namespace {

constexpr const char* kServerName = "AutoLinker Local MCP";
constexpr const char* kServerVersion = "0.0.0";
constexpr const char* kBindHost = "127.0.0.1";
constexpr int kBasePort = 19207;
constexpr int kMaxPortAttempts = 16;
constexpr std::size_t kClientWorkerCount = 4;
constexpr std::size_t kMaxQueuedClients = 32;
constexpr std::size_t kMaxExternalMcpSessions = 256;
constexpr auto kExternalMcpSessionTtl = std::chrono::minutes(30);

std::atomic_bool g_stopRequested = false;
std::atomic_bool g_running = false;
std::atomic_bool g_registryRefreshFailed = false;
std::atomic_int g_boundPort = 0;
std::mutex g_stateMutex;
std::thread g_serverThread;
SOCKET g_listenSocket = INVALID_SOCKET;
std::string g_instanceId;
std::string g_sourceFilePathHint;
std::string g_pageNameHint;
std::string g_pageTypeHint;
std::atomic_ullong g_mcpSessionCounter = 1;
std::mutex g_clientQueueMutex;
std::condition_variable g_clientQueueCv;
std::deque<SOCKET> g_clientQueue;
bool g_clientWorkersStopping = false;

struct ExternalMcpSessionState {
	bool workspaceRefreshed = false;
	std::string sourceFilePath;
	std::uint64_t mirrorGeneration = 0;
	std::chrono::steady_clock::time_point lastSeen = std::chrono::steady_clock::now();
};

std::unordered_map<std::string, ExternalMcpSessionState> g_externalMcpSessions;

struct HttpRequest {
	std::string method;
	std::string path;
	std::unordered_map<std::string, std::string> headers;
	std::string body;
};

std::string TrimAsciiCopy(const std::string& text)
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

std::string ToLowerAsciiCopy(const std::string& text)
{
	std::string lowered = text;
	std::transform(lowered.begin(), lowered.end(), lowered.begin(), [](unsigned char ch) {
		return static_cast<char>(std::tolower(ch));
	});
	return lowered;
}

void LogMcp(const std::string& message)
{
	Logger::Instance().WriteAndIde("LocalMCP", message);
}

void LogRegistryRefreshFailureOnce(const std::string& detail)
{
	if (!detail.empty() && !g_registryRefreshFailed.exchange(true)) {
		LogMcp(std::format("refresh registry failed: {}", detail));
	}
}

bool IsStrictUtf8Text(const std::string& text)
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

size_t FindValidUtf8PrefixLength(const std::string& text, size_t maxBytes)
{
	size_t prefix = (std::min)(text.size(), maxBytes);
	while (prefix > 0) {
		if (MultiByteToWideChar(
			CP_UTF8,
			MB_ERR_INVALID_CHARS,
			text.data(),
			static_cast<int>(prefix),
			nullptr,
			0) > 0) {
			return prefix;
		}
		--prefix;
	}
	return 0;
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

std::string EncodeWideToUtf8(const std::wstring& text)
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
	if (!IsStrictUtf8Text(text)) {
		return text;
	}

	const int wideLen = MultiByteToWideChar(
		CP_UTF8,
		MB_ERR_INVALID_CHARS,
		text.data(),
		static_cast<int>(text.size()),
		nullptr,
		0);
	if (wideLen <= 0) {
		return text;
	}

	std::wstring wide(static_cast<size_t>(wideLen), L'\0');
	if (MultiByteToWideChar(
		CP_UTF8,
		MB_ERR_INVALID_CHARS,
		text.data(),
		static_cast<int>(text.size()),
		wide.data(),
		wideLen) <= 0) {
		return text;
	}

	constexpr UINT kGbkCodePage = 936;
	const int gbkLen = WideCharToMultiByte(
		kGbkCodePage,
		0,
		wide.data(),
		wideLen,
		nullptr,
		0,
		nullptr,
		nullptr);
	if (gbkLen <= 0) {
		return text;
	}

	std::string gbk(static_cast<size_t>(gbkLen), '\0');
	if (WideCharToMultiByte(
		kGbkCodePage,
		0,
		wide.data(),
		wideLen,
		gbk.data(),
		gbkLen,
		nullptr,
		nullptr) <= 0) {
		return text;
	}
	return gbk;
}

void LogMcpCallSplit(const std::string& fileMessage, const std::string& ideMessage)
{
	Logger::Instance().WriteSplit("MCP", fileMessage, ideMessage);
}

std::string TrimAsciiSingleLine(const std::string& text)
{
	return TrimAsciiCopy(text);
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
	return TrimAsciiSingleLine(collapsed);
}

std::string TruncateMcpLogText(const std::string& text, size_t maxChars = 180)
{
	std::wstring wide;
	if (TryDecodeTextToWide(text, wide)) {
		if (wide.size() <= maxChars) {
			return text;
		}
		const std::wstring truncatedWide = wide.substr(0, maxChars);
		const std::string truncatedUtf8 = EncodeWideToUtf8(truncatedWide);
		if (!truncatedUtf8.empty()) {
			return truncatedUtf8 + "...";
		}
	}

	if (text.size() <= maxChars) {
		return text;
	}
	const size_t keepBytes = FindValidUtf8PrefixLength(text, maxChars);
	if (keepBytes > 0) {
		return text.substr(0, keepBytes) + "...";
	}
	return text.substr(0, maxChars) + "...";
}

std::string BuildMcpPayloadHash(const std::string& text)
{
	std::uint64_t hash = 1469598103934665603ull;
	for (const unsigned char ch : text) {
		hash ^= ch;
		hash *= 1099511628211ull;
	}
	return std::format("{:016X}", hash);
}

bool IsVerboseMcpPayloadLoggingEnabled()
{
	static const bool enabled = []() {
		char value[16] = {};
		const DWORD length = GetEnvironmentVariableA(
			"AUTOLINKER_VERBOSE_TOOL_LOG",
			value,
			static_cast<DWORD>(sizeof(value)));
		if (length == 0 || length >= sizeof(value)) {
			return false;
		}
		const std::string normalized = ToLowerAsciiCopy(TrimAsciiCopy(value));
		return normalized == "1" || normalized == "true" || normalized == "yes" || normalized == "on";
	}();
	return enabled;
}

std::string BuildMcpPayloadMetadata(const nlohmann::json& value)
{
	const std::string dumped = value.dump(-1, ' ', false, nlohmann::json::error_handler_t::replace);
	nlohmann::json metadata;
	metadata["bytes"] = dumped.size();
	metadata["hash"] = BuildMcpPayloadHash(dumped);
	if (value.is_object()) {
		static constexpr const char* kKeys[] = {
			"ok", "status", "error", "name", "file_path", "page_name", "code_hash",
			"new_hash", "verified", "count", "returned", "match_count", "has_more",
			"truncated", "isError"
		};
		for (const char* key : kKeys) {
			const auto it = value.find(key);
			if (it != value.end() && (it->is_primitive() || it->is_null())) {
				metadata[key] = *it;
			}
		}
	}
	return metadata.dump(-1, ' ', false, nlohmann::json::error_handler_t::replace);
}

std::string BuildMcpTextMetadata(const std::string& value)
{
	const nlohmann::json parsed = nlohmann::json::parse(value, nullptr, false);
	if (!parsed.is_discarded()) {
		return BuildMcpPayloadMetadata(parsed);
	}
	return nlohmann::json({
		{"bytes", value.size()},
		{"hash", BuildMcpPayloadHash(value)}
	}).dump();
}

std::string FormatMcpLogJson(const nlohmann::json& value)
{
	if (!IsVerboseMcpPayloadLoggingEnabled()) {
		return BuildMcpPayloadMetadata(value);
	}
	// 与 DumpJsonSafe 同策略：error_handler=replace，避免日志路径因非法 UTF-8 抛异常。
	return TruncateMcpLogText(
		SanitizeSingleLineText(value.dump(-1, ' ', false, nlohmann::json::error_handler_t::replace)),
		180);
}

std::string FormatMcpFileLogJson(const nlohmann::json& value)
{
	if (!IsVerboseMcpPayloadLoggingEnabled()) {
		return BuildMcpPayloadMetadata(value);
	}
	// 文件日志保留 JSON 字符串中的转义换行，避免排查时丢失原始内容。
	return value.dump(-1, ' ', false, nlohmann::json::error_handler_t::replace);
}

std::string FormatMcpLogText(const std::string& value)
{
	if (!IsVerboseMcpPayloadLoggingEnabled()) {
		return BuildMcpTextMetadata(value);
	}
	return TruncateMcpLogText(SanitizeSingleLineText(value), 180);
}

std::string FormatMcpFileLogText(const std::string& value)
{
	return IsVerboseMcpPayloadLoggingEnabled() ? value : BuildMcpTextMetadata(value);
}

std::string BuildJsonValueCallSuffix(const nlohmann::json& value)
{
	if (value.is_null()) {
		return "null";
	}
	if (value.is_object() && value.empty()) {
		return "null";
	}
	return FormatMcpLogJson(value);
}

std::string BuildJsonValueFileCallSuffix(const nlohmann::json& value)
{
	if (value.is_null()) {
		return "null";
	}
	if (value.is_object() && value.empty()) {
		return "null";
	}
	return FormatMcpFileLogJson(value);
}

struct McpLogContext {
	bool enabled = false;
	std::string responseName;
	std::string requestDisplayForIde;
	std::string requestDisplayForFile;
	std::string responseDisplayForIde;
	std::string responseDisplayForFile;
};

void SetMcpLogResponse(McpLogContext* ctx, const std::string& value)
{
	if (ctx == nullptr) {
		return;
	}
	ctx->responseDisplayForIde = value;
	ctx->responseDisplayForFile = value;
}

void SetMcpLogResponseJson(McpLogContext* ctx, const nlohmann::json& value)
{
	if (ctx == nullptr) {
		return;
	}
	ctx->responseDisplayForIde = FormatMcpLogJson(value);
	ctx->responseDisplayForFile = FormatMcpFileLogJson(value);
}

McpLogContext BuildMcpLogContextForPayload(const nlohmann::json& payload)
{
	McpLogContext ctx{};
	ctx.enabled = true;

	if (!payload.is_object()) {
		ctx.responseName = "invalid_request";
		ctx.requestDisplayForIde = "invalid_request(" + FormatMcpLogJson(payload) + ")";
		ctx.requestDisplayForFile = "invalid_request(" + FormatMcpFileLogJson(payload) + ")";
		return ctx;
	}

	const std::string method = payload.contains("method") && payload["method"].is_string()
		? payload["method"].get<std::string>()
		: std::string("unknown_method");
	const nlohmann::json params = payload.contains("params") ? payload["params"] : nlohmann::json(nullptr);

	if (method == "tools/call" &&
		params.is_object() &&
		params.contains("name") &&
		params["name"].is_string()) {
		const std::string toolName = params["name"].get<std::string>();
		const nlohmann::json arguments = params.contains("arguments")
			? params["arguments"]
			: nlohmann::json(nullptr);
		ctx.responseName = toolName;
		ctx.requestDisplayForIde = toolName + "(" + BuildJsonValueCallSuffix(arguments) + ")";
		ctx.requestDisplayForFile = toolName + "(" + BuildJsonValueFileCallSuffix(arguments) + ")";
		return ctx;
	}

	ctx.responseName = method;
	ctx.requestDisplayForIde = method + "(" + BuildJsonValueCallSuffix(params) + ")";
	ctx.requestDisplayForFile = method + "(" + BuildJsonValueFileCallSuffix(params) + ")";
	return ctx;
}

McpLogContext BuildMcpLogContextForRequestBody(const std::string& body)
{
	try {
		const nlohmann::json payload = body.empty()
			? nlohmann::json::object()
			: nlohmann::json::parse(body);
		return BuildMcpLogContextForPayload(payload);
	}
	catch (...) {
		McpLogContext ctx{};
		ctx.enabled = true;
		ctx.responseName = "invalid_json";
		ctx.requestDisplayForIde = "invalid_json(" + FormatMcpLogText(body) + ")";
		ctx.requestDisplayForFile = "invalid_json(" + FormatMcpFileLogText(body) + ")";
		return ctx;
	}
}

void LogMcpRequest(const McpLogContext& ctx)
{
	if (!ctx.enabled) {
		return;
	}
	const std::string ideText = ctx.requestDisplayForIde.empty() ? "null" : ctx.requestDisplayForIde;
	const std::string fileText = ctx.requestDisplayForFile.empty() ? ideText : ctx.requestDisplayForFile;
	LogMcpCallSplit(">> " + fileText, ">> " + ideText);
}

void LogMcpResponse(const McpLogContext& ctx, double elapsedMs)
{
	if (!ctx.enabled) {
		return;
	}

	const std::string responseTextForIde = ctx.responseDisplayForIde.empty() ? "null" : ctx.responseDisplayForIde;
	const std::string responseTextForFile = ctx.responseDisplayForFile.empty() ? responseTextForIde : ctx.responseDisplayForFile;
	LogMcpCallSplit(std::format(
		"<< {} ({:.1f}ms) {}",
		ctx.responseName.empty() ? "unknown" : ctx.responseName,
		elapsedMs,
		responseTextForFile),
		std::format(
		"<< {} ({:.1f}ms) {}",
		ctx.responseName.empty() ? "unknown" : ctx.responseName,
		elapsedMs,
		responseTextForIde));
}

void CloseSocketSafe(SOCKET& sock)
{
	if (sock == INVALID_SOCKET) {
		return;
	}
	shutdown(sock, SD_BOTH);
	closesocket(sock);
	sock = INVALID_SOCKET;
}

bool IsPortInUseError(int error)
{
	return error == WSAEADDRINUSE || error == WSAEACCES;
}

std::string GetCurrentProcessPathLocal()
{
	char buffer[MAX_PATH] = {};
	const DWORD len = GetModuleFileNameA(nullptr, buffer, static_cast<DWORD>(sizeof(buffer)));
	if (len == 0 || len >= sizeof(buffer)) {
		return std::string();
	}
	return std::string(buffer, buffer + len);
}

std::string GetCurrentProcessNameLocal()
{
	const std::string fullPath = GetCurrentProcessPathLocal();
	if (fullPath.empty()) {
		return std::string();
	}
	const size_t pos = fullPath.find_last_of("\\/");
	return pos == std::string::npos ? fullPath : fullPath.substr(pos + 1);
}

std::string BuildEndpointForPort(int port)
{
	if (port <= 0) {
		return std::string();
	}
	return std::format("http://{}:{}/mcp", kBindHost, port);
}

std::string GenerateInstanceId()
{
	return std::format("pid-{}-{:X}", GetCurrentProcessId(), GetTickCount64());
}

bool IsValidMcpSessionId(const std::string& sessionId)
{
	if (sessionId.empty() || sessionId.size() > 128) {
		return false;
	}
	for (const unsigned char ch : sessionId) {
		if (std::isalnum(ch) == 0 && ch != '-' && ch != '_' && ch != '.') {
			return false;
		}
	}
	return true;
}

std::string GenerateMcpSessionId()
{
	return std::format(
		"mcp-{}-{:X}-{}",
		GetCurrentProcessId(),
		GetTickCount64(),
		g_mcpSessionCounter.fetch_add(1));
}

bool IsExternalWorkspaceRefreshed(const std::string& sessionId)
{
	std::lock_guard<std::mutex> lock(g_stateMutex);
	const auto it = g_externalMcpSessions.find(sessionId);
	if (it == g_externalMcpSessions.end()) {
		return false;
	}
	it->second.lastSeen = std::chrono::steady_clock::now();
	return it->second.workspaceRefreshed &&
		it->second.sourceFilePath == g_sourceFilePathHint &&
		it->second.mirrorGeneration == WorkspaceMirror::GetGeneration();
}

void PruneExternalMcpSessionsLocked(std::chrono::steady_clock::time_point now)
{
	for (auto it = g_externalMcpSessions.begin(); it != g_externalMcpSessions.end();) {
		if (it->first != "legacy" && now - it->second.lastSeen >= kExternalMcpSessionTtl) {
			ClearToolApprovalScope("external-mcp:" + it->first);
			it = g_externalMcpSessions.erase(it);
		}
		else {
			++it;
		}
	}
	while (g_externalMcpSessions.size() >= kMaxExternalMcpSessions) {
		auto oldest = g_externalMcpSessions.end();
		for (auto it = g_externalMcpSessions.begin(); it != g_externalMcpSessions.end(); ++it) {
			if (it->first == "legacy") {
				continue;
			}
			if (oldest == g_externalMcpSessions.end() || it->second.lastSeen < oldest->second.lastSeen) {
				oldest = it;
			}
		}
		if (oldest == g_externalMcpSessions.end()) {
			break;
		}
		ClearToolApprovalScope("external-mcp:" + oldest->first);
		g_externalMcpSessions.erase(oldest);
	}
}

void SetExternalWorkspaceRefreshed(const std::string& sessionId, bool refreshed)
{
	std::lock_guard<std::mutex> lock(g_stateMutex);
	const auto now = std::chrono::steady_clock::now();
	PruneExternalMcpSessionsLocked(now);
	ExternalMcpSessionState& state = g_externalMcpSessions[sessionId];
	state.workspaceRefreshed = refreshed;
	state.sourceFilePath = refreshed ? g_sourceFilePathHint : std::string();
	state.mirrorGeneration = refreshed ? WorkspaceMirror::GetGeneration() : 0;
	state.lastSeen = now;
}

void RemoveExternalMcpSession(const std::string& sessionId)
{
	if (sessionId.empty() || sessionId == "legacy") {
		return;
	}
	std::lock_guard<std::mutex> lock(g_stateMutex);
	ClearToolApprovalScope("external-mcp:" + sessionId);
	g_externalMcpSessions.erase(sessionId);
}

void ResetExternalWorkspaceRefreshStatesLocked()
{
	for (auto& [sessionId, state] : g_externalMcpSessions) {
		(void)sessionId;
		state.workspaceRefreshed = false;
		state.sourceFilePath.clear();
		state.mirrorGeneration = 0;
		state.lastSeen = std::chrono::steady_clock::now();
	}
}

LocalMcpInstanceRegistry::InstanceRecord BuildCurrentInstanceRecord()
{
	LocalMcpInstanceRegistry::InstanceRecord record;
	record.instanceId = g_instanceId;
	record.processId = GetCurrentProcessId();
	record.processPath = GetCurrentProcessPathLocal();
	record.processName = GetCurrentProcessNameLocal();
	record.port = g_boundPort.load();
	record.endpoint = BuildEndpointForPort(record.port);
	{
		std::lock_guard<std::mutex> lock(g_stateMutex);
		record.sourceFilePathHint = g_sourceFilePathHint;
		record.pageNameHint = g_pageNameHint;
		record.pageTypeHint = g_pageTypeHint;
	}
	record.lastSeenUnixMs = static_cast<std::uint64_t>(
		std::chrono::duration_cast<std::chrono::milliseconds>(
			std::chrono::system_clock::now().time_since_epoch()).count());
	return record;
}

void RefreshCurrentInstanceRegistry()
{
	try {
		const LocalMcpInstanceRegistry::InstanceRecord record = BuildCurrentInstanceRecord();
		if (record.instanceId.empty() || record.port <= 0 || record.endpoint.empty()) {
			return;
		}
		std::string error;
		if (!LocalMcpInstanceRegistry::UpsertCurrentInstance(record, &error)) {
			LogRegistryRefreshFailureOnce(error);
		}
		else if (g_registryRefreshFailed.exchange(false)) {
			LogMcp("instance registry refresh recovered");
		}
	}
	catch (const std::exception& ex) {
		LogRegistryRefreshFailureOnce(std::format("exception: {}", ex.what()));
	}
	catch (...) {
		LogRegistryRefreshFailureOnce("exception: unknown");
	}
}

void RemoveCurrentInstanceRegistry()
{
	try {
		if (g_instanceId.empty()) {
			return;
		}
		std::string error;
		if (!LocalMcpInstanceRegistry::RemoveCurrentInstance(g_instanceId, &error) && !error.empty()) {
			LogMcp(std::format("remove registry failed: {}", error));
		}
	}
	catch (const std::exception& ex) {
		LogMcp(std::format("remove registry exception: {}", ex.what()));
	}
	catch (...) {
		LogMcp("remove registry exception: unknown");
	}
}

bool ReadExactBytes(SOCKET sock, std::string& buffer, size_t wantedBytes)
{
	while (buffer.size() < wantedBytes) {
		char temp[4096];
		const int toRead = static_cast<int>((std::min)(wantedBytes - buffer.size(), sizeof(temp)));
		const int received = recv(sock, temp, toRead, 0);
		if (received <= 0) {
			return false;
		}
		buffer.append(temp, static_cast<size_t>(received));
	}
	return true;
}

bool ReadHttpRequest(SOCKET sock, HttpRequest& outRequest)
{
	outRequest = {};
	std::string raw;
	size_t headerEnd = std::string::npos;
	while ((headerEnd = raw.find("\r\n\r\n")) == std::string::npos) {
		char temp[4096];
		const int received = recv(sock, temp, static_cast<int>(sizeof(temp)), 0);
		if (received <= 0) {
			return false;
		}
		raw.append(temp, static_cast<size_t>(received));
		if (raw.size() > 1024 * 1024) {
			return false;
		}
	}

	const std::string headerText = raw.substr(0, headerEnd);
	std::string remaining = raw.substr(headerEnd + 4);

	size_t lineBegin = 0;
	const size_t requestLineEnd = headerText.find("\r\n");
	const std::string requestLine = requestLineEnd == std::string::npos
		? headerText
		: headerText.substr(0, requestLineEnd);
	lineBegin = requestLineEnd == std::string::npos ? headerText.size() : requestLineEnd + 2;

	const size_t firstSpace = requestLine.find(' ');
	const size_t secondSpace = firstSpace == std::string::npos ? std::string::npos : requestLine.find(' ', firstSpace + 1);
	if (firstSpace == std::string::npos || secondSpace == std::string::npos) {
		return false;
	}
	outRequest.method = requestLine.substr(0, firstSpace);
	outRequest.path = requestLine.substr(firstSpace + 1, secondSpace - firstSpace - 1);

	while (lineBegin < headerText.size()) {
		const size_t lineEnd = headerText.find("\r\n", lineBegin);
		const std::string line = lineEnd == std::string::npos
			? headerText.substr(lineBegin)
			: headerText.substr(lineBegin, lineEnd - lineBegin);
		lineBegin = lineEnd == std::string::npos ? headerText.size() : lineEnd + 2;
		if (line.empty()) {
			continue;
		}
		const size_t colon = line.find(':');
		if (colon == std::string::npos) {
			continue;
		}
		const std::string key = ToLowerAsciiCopy(TrimAsciiCopy(line.substr(0, colon)));
		const std::string value = TrimAsciiCopy(line.substr(colon + 1));
		outRequest.headers[key] = value;
	}

	size_t contentLength = 0;
	const auto contentIt = outRequest.headers.find("content-length");
	if (contentIt != outRequest.headers.end()) {
		contentLength = static_cast<size_t>(std::strtoul(contentIt->second.c_str(), nullptr, 10));
	}

	if (contentLength > 2 * 1024 * 1024) {
		return false;
	}

	if (!ReadExactBytes(sock, remaining, contentLength)) {
		return false;
	}
	outRequest.body = remaining.substr(0, contentLength);
	return true;
}

std::string BuildHttpResponse(
	int statusCode,
	const char* statusText,
	const char* contentType,
	const std::string& body,
	const std::string& extraHeaders = std::string())
{
	return std::format(
		"HTTP/1.1 {} {}\r\n"
		"Content-Type: {}\r\n"
		"Content-Length: {}\r\n"
		"Connection: close\r\n"
		"Cache-Control: no-store\r\n"
		"{}"
		"\r\n{}",
		statusCode,
		statusText,
		contentType,
		body.size(),
		extraHeaders,
		body);
}

void SendHttpResponse(
	SOCKET sock,
	int statusCode,
	const char* statusText,
	const char* contentType,
	const std::string& body,
	const std::string& extraHeaders = std::string())
{
	const std::string response = BuildHttpResponse(statusCode, statusText, contentType, body, extraHeaders);
	size_t sentTotal = 0;
	while (sentTotal < response.size()) {
		const int sent = send(
			sock,
			response.data() + sentTotal,
			static_cast<int>(response.size() - sentTotal),
			0);
		if (sent <= 0) {
			return;
		}
		sentTotal += static_cast<size_t>(sent);
	}
}

// 安全序列化 JSON 为字符串：ensure_ascii=false 保留中文可读，error_handler=replace 遇非法 UTF-8
// 用 U+FFFD 替换而非抛 type_error.316，避免外部非法输入经 .dump() 抛异常导致线程/进程崩溃。
std::string DumpJsonSafe(const nlohmann::json& value)
{
	return value.dump(-1, ' ', false, nlohmann::json::error_handler_t::replace);
}

nlohmann::json BuildJsonRpcError(const nlohmann::json& id, int code, const std::string& message)
{
	return {
		{"jsonrpc", "2.0"},
		{"id", id},
		{"error", {
			{"code", code},
			{"message", message}
		}}
	};
}

nlohmann::json BuildJsonRpcResult(const nlohmann::json& id, const nlohmann::json& result)
{
	return {
		{"jsonrpc", "2.0"},
		{"id", id},
		{"result", result}
	};
}

std::string NegotiateProtocolVersion(const nlohmann::json& params)
{
	const std::string requested = params.is_object()
		? params.value("protocolVersion", std::string())
		: std::string();
	if (requested == "2025-11-25" ||
		requested == "2025-03-26" ||
		requested == "2024-11-05") {
		return requested;
	}
	return "2025-11-25";
}

nlohmann::json BuildInitializeResult(const nlohmann::json& params)
{
	return {
		{"protocolVersion", NegotiateProtocolVersion(params)},
		{"capabilities", {
			{"tools", {
				{"listChanged", false}
			}}
		}},
		{"serverInfo", {
			{"name", kServerName},
			{"version", kServerVersion}
		}}
	};
}

bool TryBuildToolListResult(nlohmann::json& outResult, std::string& outError)
{
	try {
		const nlohmann::json tools = nlohmann::json::parse(AIService::BuildPublicToolCatalogJson());
		if (!tools.is_array()) {
			outError = "public tool catalog must be an array";
			return false;
		}
		outResult = {{"tools", tools}};
		return true;
	}
	catch (const std::exception& ex) {
		outError = std::string("build tools list failed: ") + ex.what();
		return false;
	}
}

bool TryFindExternalToolDefinition(
	const std::string& toolName,
	nlohmann::json& outDefinition,
	std::string& outError)
{
	outDefinition = nlohmann::json::object();
	outError.clear();
	if (!AIChatToolRegistry::IsExternalPublic(toolName)) {
		outError = "unknown or unavailable external tool: " + toolName;
		return false;
	}

	try {
		const nlohmann::json catalog = nlohmann::json::parse(AIService::BuildPublicToolCatalogJson());
		if (!catalog.is_array()) {
			outError = "public tool catalog must be an array";
			return false;
		}
		for (const auto& item : catalog) {
			if (item.is_object() && item.value("name", std::string()) == toolName) {
				outDefinition = item;
				return true;
			}
		}
		outError = "tool is disabled by the current AutoLinker configuration: " + toolName;
		return false;
	}
	catch (const std::exception& ex) {
		outError = std::string("build public tool catalog failed: ") + ex.what();
		return false;
	}
}

std::string BuildToolContentSummary(
	const std::string& toolName,
	const nlohmann::json& structured,
	bool structuredOk,
	const std::string& rawResult)
{
	if (!structuredOk) {
		return TruncateMcpLogText(rawResult, 2048);
	}

	nlohmann::json summary;
	summary["tool"] = toolName;
	summary["note"] = "Full machine-readable result is available in structuredContent.";
	if (structured.is_object()) {
		static constexpr const char* kSummaryKeys[] = {
			"ok", "status", "error", "file_path", "page_name", "mapped_page_name",
			"code_hash", "new_hash", "verified", "count", "returned", "returned_lines",
			"match_count", "files_with_matches", "has_more", "next_offset", "truncated"
		};
		for (const char* key : kSummaryKeys) {
			const auto it = structured.find(key);
			if (it != structured.end() && (it->is_primitive() || it->is_null())) {
				summary[key] = *it;
			}
		}
	}
	return DumpJsonSafe(summary);
}

bool TryBuildToolCallResult(
	const nlohmann::json& params,
	const std::string& sessionId,
	nlohmann::json& outResult,
	std::string& outError)
{
	if (!params.is_object()) {
		outError = "tools/call params must be object";
		return false;
	}
	if (!params.contains("name") || !params["name"].is_string()) {
		outError = "tools/call requires string params.name";
		return false;
	}

	const std::string toolName = params["name"].get<std::string>();
	nlohmann::json arguments = nlohmann::json::object();
	if (params.contains("arguments") && !params["arguments"].is_null()) {
		arguments = params["arguments"];
	}
	if (!arguments.is_object()) {
		outError = "tools/call params.arguments must be object";
		return false;
	}

	nlohmann::json toolDefinition;
	if (!TryFindExternalToolDefinition(toolName, toolDefinition, outError)) {
		return false;
	}
	const nlohmann::json inputSchema = toolDefinition.value(
		"inputSchema",
		nlohmann::json::object({
			{"type", "object"},
			{"properties", nlohmann::json::object()},
			{"additionalProperties", false}
		}));
	std::string validationError;
	if (!AIChatToolRegistry::ValidateArguments(arguments, inputSchema, validationError)) {
		outError = "invalid tool arguments: " + validationError;
		return false;
	}
	const bool requiresBaseHash =
		toolName == "edit_file" ||
		toolName == "multi_edit_file" ||
		toolName == "write_file" ||
		toolName == "diff_file";
	if (requiresBaseHash &&
		(!arguments.contains("expected_base_hash") ||
			!arguments["expected_base_hash"].is_string() ||
			TrimAsciiCopy(arguments["expected_base_hash"].get<std::string>()).empty())) {
		outError = "external source edits and diffs require expected_base_hash from read_real_file";
		return false;
	}
	if (toolName == "restore_file_snapshot" &&
		(!arguments.contains("expected_current_hash") ||
			!arguments["expected_current_hash"].is_string() ||
			TrimAsciiCopy(arguments["expected_current_hash"].get<std::string>()).empty())) {
		outError = "external snapshot restore requires expected_current_hash from read_real_file";
		return false;
	}

	std::string resultJsonUtf8;
	bool toolOk = false;
	if (AIChatToolRegistry::RequiresWorkspaceRefresh(toolName) && !IsExternalWorkspaceRefreshed(sessionId)) {
		resultJsonUtf8 = nlohmann::json({
			{"ok", false},
			{"error", "workspace_refresh_required"},
			{"required_tool", "refresh_workspace_mirror"},
			{"hint", "External MCP clients must call refresh_workspace_mirror successfully before source read/edit tools in each MCP session."}
		}).dump();
	}
	else {
		if (!AIChatFeature::ExecutePublicTool(
				toolName,
				arguments.dump(),
				resultJsonUtf8,
				toolOk,
				[]() { return g_stopRequested.load(); },
				"external-mcp:" + sessionId)) {
			outError = "tool execution transport failed";
			return false;
		}
		if (toolOk) {
			const AIChatToolRegistry::ToolMetadata* metadata = AIChatToolRegistry::Find(toolName);
			if (toolName == "refresh_workspace_mirror" ||
				toolName == "add_new_file" ||
				(metadata != nullptr && metadata->requiresWorkspaceRefresh)) {
				SetExternalWorkspaceRefreshed(sessionId, true);
			}
		}
	}

	bool isError = !toolOk;
	nlohmann::json structured;
	bool structuredOk = false;
	try {
		structured = nlohmann::json::parse(resultJsonUtf8);
		structuredOk = true;
	}
	catch (...) {
		structuredOk = false;
	}

	if (structuredOk && structured.is_object() && structured.contains("ok") && structured["ok"].is_boolean()) {
		isError = !structured["ok"].get<bool>();
	}

	const std::string contentSummary = BuildToolContentSummary(
		toolName,
		structured,
		structuredOk,
		resultJsonUtf8);
	outResult = {
		{"content", nlohmann::json::array({
			{
				{"type", "text"},
				{"text", contentSummary}
			}
		})},
		{"isError", isError}
	};
	if (structuredOk) {
		outResult["structuredContent"] = structured;
	}
	return true;
}

bool TryHandleJsonRpc(
	const HttpRequest& request,
	int& outStatusCode,
	std::string& outBody,
	std::string& outResponseSessionId,
	McpLogContext* outLogContext)
{
	outStatusCode = 200;
	outBody.clear();
	outResponseSessionId.clear();
	if (outLogContext != nullptr) {
		*outLogContext = {};
	}

	nlohmann::json payload;
	try {
		payload = request.body.empty() ? nlohmann::json::object() : nlohmann::json::parse(request.body);
	}
	catch (const std::exception& ex) {
		if (outLogContext != nullptr) {
			outLogContext->enabled = true;
			outLogContext->responseName = "invalid_json";
			outLogContext->requestDisplayForIde = "invalid_json(" + FormatMcpLogText(request.body) + ")";
			outLogContext->requestDisplayForFile = "invalid_json(" + FormatMcpFileLogText(request.body) + ")";
			SetMcpLogResponseJson(outLogContext, {
				{"code", -32700},
				{"message", std::string("parse error: ") + ex.what()}
			});
		}
		outBody = DumpJsonSafe(BuildJsonRpcError(nullptr, -32700, std::string("parse error: ") + ex.what()));
		return true;
	}

	if (outLogContext != nullptr) {
		*outLogContext = BuildMcpLogContextForPayload(payload);
	}

	if (!payload.is_object()) {
		SetMcpLogResponseJson(outLogContext, {
			{"code", -32600},
			{"message", "request must be a JSON object"}
		});
		outBody = DumpJsonSafe(BuildJsonRpcError(nullptr, -32600, "request must be a JSON object"));
		return true;
	}

	const nlohmann::json id = payload.contains("id") ? payload["id"] : nlohmann::json(nullptr);
	if (!payload.contains("method") || !payload["method"].is_string()) {
		SetMcpLogResponseJson(outLogContext, {
			{"code", -32600},
			{"message", "method is required"}
		});
		outBody = DumpJsonSafe(BuildJsonRpcError(id, -32600, "method is required"));
		return true;
	}

	const std::string method = payload["method"].get<std::string>();
	const bool hasId = payload.contains("id");
	const nlohmann::json params = payload.contains("params") ? payload["params"] : nlohmann::json::object();
	std::string requestSessionId = "legacy";
	if (const auto it = request.headers.find("mcp-session-id");
		it != request.headers.end() && IsValidMcpSessionId(it->second)) {
		requestSessionId = it->second;
	}

	if (method == "notifications/initialized") {
		outStatusCode = 202;
		outBody.clear();
		SetMcpLogResponse(outLogContext, "null");
		return true;
	}

	if (method == "ping") {
		const nlohmann::json result = nlohmann::json::object();
		SetMcpLogResponseJson(outLogContext, result);
		outBody = DumpJsonSafe(BuildJsonRpcResult(id, result));
		return true;
	}

	if (method == "initialize") {
		outResponseSessionId = requestSessionId == "legacy" ? GenerateMcpSessionId() : requestSessionId;
		{
			std::lock_guard<std::mutex> lock(g_stateMutex);
			const auto now = std::chrono::steady_clock::now();
			PruneExternalMcpSessionsLocked(now);
			ClearToolApprovalScope("external-mcp:" + outResponseSessionId);
			ExternalMcpSessionState& state = g_externalMcpSessions[outResponseSessionId];
			state.workspaceRefreshed = false;
			state.sourceFilePath.clear();
			state.mirrorGeneration = 0;
			state.lastSeen = now;
			if (requestSessionId == "legacy") {
				ExternalMcpSessionState& legacyState = g_externalMcpSessions["legacy"];
				legacyState.workspaceRefreshed = false;
				legacyState.sourceFilePath.clear();
				legacyState.mirrorGeneration = 0;
				legacyState.lastSeen = now;
			}
		}
		const nlohmann::json result = BuildInitializeResult(params);
		SetMcpLogResponseJson(outLogContext, result);
		outBody = DumpJsonSafe(BuildJsonRpcResult(id, result));
		return true;
	}

	if (method == "tools/list") {
		nlohmann::json result;
		std::string error;
		if (!TryBuildToolListResult(result, error)) {
			SetMcpLogResponseJson(outLogContext, {
				{"code", -32603},
				{"message", error}
			});
			outBody = DumpJsonSafe(BuildJsonRpcError(id, -32603, error));
			return true;
		}
		SetMcpLogResponseJson(outLogContext, result);
		outBody = DumpJsonSafe(BuildJsonRpcResult(id, result));
		return true;
	}

	if (method == "tools/call") {
		nlohmann::json result;
		std::string error;
		if (!TryBuildToolCallResult(params, requestSessionId, result, error)) {
			SetMcpLogResponseJson(outLogContext, {
				{"code", -32602},
				{"message", error}
			});
			outBody = DumpJsonSafe(BuildJsonRpcError(id, -32602, error));
			return true;
		}
		SetMcpLogResponseJson(outLogContext, result);
		outBody = DumpJsonSafe(BuildJsonRpcResult(id, result));
		return true;
	}

	if (!hasId) {
		outStatusCode = 202;
		outBody.clear();
		SetMcpLogResponse(outLogContext, "null");
		return true;
	}

	SetMcpLogResponseJson(outLogContext, {
		{"code", -32601},
		{"message", "method not found"}
	});
	outBody = DumpJsonSafe(BuildJsonRpcError(id, -32601, "method not found"));
	return true;
}

void HandleClientImpl(SOCKET clientSock)
{
	DWORD timeoutMs = 5000;
	setsockopt(clientSock, SOL_SOCKET, SO_RCVTIMEO, reinterpret_cast<const char*>(&timeoutMs), sizeof(timeoutMs));
	setsockopt(clientSock, SOL_SOCKET, SO_SNDTIMEO, reinterpret_cast<const char*>(&timeoutMs), sizeof(timeoutMs));

	HttpRequest request;
	if (!ReadHttpRequest(clientSock, request)) {
		SendHttpResponse(clientSock, 400, "Bad Request", "application/json; charset=utf-8", R"({"ok":false,"error":"invalid http request"})");
		return;
	}

	// 本地 MCP 只服务原生桌面客户端。浏览器请求会携带 Origin；若允许通配 CORS，
	// 任意网页都能读取或修改当前 IDE 工程，因此在尚未配置显式浏览器能力令牌前拒绝此类请求。
	if (const auto originIt = request.headers.find("origin");
		originIt != request.headers.end() && !TrimAsciiCopy(originIt->second).empty()) {
		SendHttpResponse(clientSock, 403, "Forbidden", "application/json; charset=utf-8",
			R"({"ok":false,"error":"browser_origin_not_allowed"})");
		return;
	}

	if (request.path != "/" && request.path != "/mcp") {
		SendHttpResponse(clientSock, 404, "Not Found", "application/json; charset=utf-8", R"({"ok":false,"error":"not found"})");
		return;
	}

	if (request.method == "OPTIONS") {
		SendHttpResponse(clientSock, 204, "No Content", "text/plain; charset=utf-8", "");
		return;
	}

	if (request.method == "GET") {
		nlohmann::json health = {
			{"ok", true},
			{"service", kServerName},
			{"version", kServerVersion},
			{"instance_id", g_instanceId},
			{"process_id", GetCurrentProcessId()},
			{"port", g_boundPort.load()},
			{"mcp_endpoint", BuildEndpointForPort(g_boundPort.load())}
		};
		SendHttpResponse(clientSock, 200, "OK", "application/json; charset=utf-8", DumpJsonSafe(health));
		return;
	}

	if (request.method == "DELETE") {
		const auto sessionIt = request.headers.find("mcp-session-id");
		if (sessionIt != request.headers.end() && IsValidMcpSessionId(sessionIt->second)) {
			RemoveExternalMcpSession(sessionIt->second);
		}
		SendHttpResponse(clientSock, 204, "No Content", "text/plain; charset=utf-8", "");
		return;
	}

	if (request.method != "POST") {
		SendHttpResponse(clientSock, 405, "Method Not Allowed", "application/json; charset=utf-8", R"({"ok":false,"error":"method not allowed"})");
		return;
	}

	McpLogContext logContext = BuildMcpLogContextForRequestBody(request.body);
	LogMcpRequest(logContext);
	const auto startTime = std::chrono::steady_clock::now();

	int statusCode = 200;
	std::string responseBody;
	std::string responseSessionId;
	McpLogContext handledLogContext;
	if (!TryHandleJsonRpc(request, statusCode, responseBody, responseSessionId, &handledLogContext)) {
		const double elapsedMs = std::chrono::duration<double, std::milli>(
			std::chrono::steady_clock::now() - startTime).count();
		logContext.responseName = handledLogContext.responseName.empty() ? logContext.responseName : handledLogContext.responseName;
		logContext.responseDisplayForIde = R"({"ok":false,"error":"internal server error"})";
		logContext.responseDisplayForFile = R"({"ok":false,"error":"internal server error"})";
		LogMcpResponse(logContext, elapsedMs);
		SendHttpResponse(clientSock, 500, "Internal Server Error", "application/json; charset=utf-8", R"({"ok":false,"error":"internal server error"})");
		return;
	}
	if (handledLogContext.enabled) {
		if (!handledLogContext.responseName.empty()) {
			logContext.responseName = handledLogContext.responseName;
		}
		if (!handledLogContext.requestDisplayForIde.empty()) {
			logContext.requestDisplayForIde = handledLogContext.requestDisplayForIde;
		}
		if (!handledLogContext.requestDisplayForFile.empty()) {
			logContext.requestDisplayForFile = handledLogContext.requestDisplayForFile;
		}
		logContext.responseDisplayForIde = handledLogContext.responseDisplayForIde;
		logContext.responseDisplayForFile = handledLogContext.responseDisplayForFile;
	}

	const char* statusText = "OK";
	if (statusCode == 202) {
		statusText = "Accepted";
	}
	else if (statusCode == 204) {
		statusText = "No Content";
	}
	const double elapsedMs = std::chrono::duration<double, std::milli>(
		std::chrono::steady_clock::now() - startTime).count();
	LogMcpResponse(logContext, elapsedMs);
	const std::string sessionHeader = responseSessionId.empty()
		? std::string()
		: "Mcp-Session-Id: " + responseSessionId + "\r\n";
	SendHttpResponse(
		clientSock,
		statusCode,
		statusText,
		"application/json; charset=utf-8",
		responseBody,
		sessionHeader);
}

// 防御纵深：单个连接的任何未捕获异常（如非法 UTF-8 经 .dump() 抛 type_error）不得逃出到
// accept 线程而使整个宿主进程崩溃。此处兜底，返回 500 并保证连接被正常关闭。
void HandleClient(SOCKET clientSock)
{
	try {
		HandleClientImpl(clientSock);
	}
	catch (const std::exception& ex) {
		LogMcp(std::format("HandleClient exception: {}", ex.what()));
		try {
			SendHttpResponse(clientSock, 500, "Internal Server Error",
				"application/json; charset=utf-8", R"({"ok":false,"error":"internal server error"})");
		}
		catch (...) {
		}
	}
	catch (...) {
		LogMcp("HandleClient unknown exception");
		try {
			SendHttpResponse(clientSock, 500, "Internal Server Error",
				"application/json; charset=utf-8", R"({"ok":false,"error":"internal server error"})");
		}
		catch (...) {
		}
	}
}

void ClientWorkerMain()
{
	for (;;) {
		SOCKET clientSock = INVALID_SOCKET;
		{
			std::unique_lock<std::mutex> lock(g_clientQueueMutex);
			g_clientQueueCv.wait(lock, []() {
				return g_clientWorkersStopping || !g_clientQueue.empty();
			});
			if (g_clientQueue.empty()) {
				if (g_clientWorkersStopping) {
					return;
				}
				continue;
			}
			clientSock = g_clientQueue.front();
			g_clientQueue.pop_front();
		}

		HandleClient(clientSock);
		CloseSocketSafe(clientSock);
	}
}

bool TryQueueClient(SOCKET clientSock)
{
	std::lock_guard<std::mutex> lock(g_clientQueueMutex);
	if (g_clientWorkersStopping || g_clientQueue.size() >= kMaxQueuedClients) {
		return false;
	}
	g_clientQueue.push_back(clientSock);
	g_clientQueueCv.notify_one();
	return true;
}

void StopClientWorkersAndCloseQueue()
{
	{
		std::lock_guard<std::mutex> lock(g_clientQueueMutex);
		g_clientWorkersStopping = true;
		for (SOCKET& clientSock : g_clientQueue) {
			CloseSocketSafe(clientSock);
		}
		g_clientQueue.clear();
	}
	g_clientQueueCv.notify_all();
}

bool TryCreateListeningSocket(int& outPort)
{
	outPort = 0;
	for (int port = kBasePort; port < kBasePort + kMaxPortAttempts; ++port) {
		SOCKET listenSock = socket(AF_INET, SOCK_STREAM, IPPROTO_TCP);
		if (listenSock == INVALID_SOCKET) {
			continue;
		}

		BOOL exclusive = TRUE;
		setsockopt(listenSock, SOL_SOCKET, SO_EXCLUSIVEADDRUSE, reinterpret_cast<const char*>(&exclusive), sizeof(exclusive));

		sockaddr_in addr = {};
		addr.sin_family = AF_INET;
		addr.sin_port = htons(static_cast<u_short>(port));
		if (inet_pton(AF_INET, kBindHost, &addr.sin_addr) != 1) {
			CloseSocketSafe(listenSock);
			return false;
		}

		if (bind(listenSock, reinterpret_cast<const sockaddr*>(&addr), sizeof(addr)) == SOCKET_ERROR) {
			const int error = WSAGetLastError();
			CloseSocketSafe(listenSock);
			if (IsPortInUseError(error)) {
				continue;
			}
			LogMcp(std::format("bind {}:{} failed, error={}", kBindHost, port, error));
			return false;
		}

		if (listen(listenSock, SOMAXCONN) == SOCKET_ERROR) {
			const int error = WSAGetLastError();
			CloseSocketSafe(listenSock);
			LogMcp(std::format("listen {}:{} failed, error={}", kBindHost, port, error));
			return false;
		}

		{
			std::lock_guard<std::mutex> lock(g_stateMutex);
			g_listenSocket = listenSock;
		}
		outPort = port;
		return true;
	}
	return false;
}

void ServerThreadMain()
{
	WSADATA wsaData = {};
	if (WSAStartup(MAKEWORD(2, 2), &wsaData) != 0) {
		LogMcp("WSAStartup failed");
		return;
	}

	int boundPort = 0;
	if (!TryCreateListeningSocket(boundPort)) {
		LogMcp(std::format("failed to bind {} starting at port {}", kBindHost, kBasePort));
		WSACleanup();
		return;
	}

	g_boundPort.store(boundPort);
	{
		std::lock_guard<std::mutex> lock(g_clientQueueMutex);
		g_clientWorkersStopping = false;
		g_clientQueue.clear();
	}
	std::vector<std::thread> clientWorkers;
	try {
		clientWorkers.reserve(kClientWorkerCount);
		for (std::size_t i = 0; i < kClientWorkerCount; ++i) {
			clientWorkers.emplace_back(ClientWorkerMain);
		}
	}
	catch (const std::exception& ex) {
		LogMcp(std::format("start client workers failed: {}", ex.what()));
		StopClientWorkersAndCloseQueue();
		for (std::thread& worker : clientWorkers) {
			if (worker.joinable()) {
				worker.join();
			}
		}
		{
			std::lock_guard<std::mutex> lock(g_stateMutex);
			CloseSocketSafe(g_listenSocket);
		}
		g_boundPort.store(0);
		WSACleanup();
		return;
	}
	g_running.store(true);
	LogMcp(std::format(
		"本地 MCP 服务已启动：http://{}:{}/mcp",
		kBindHost,
		boundPort));
	RefreshCurrentInstanceRegistry();
	auto lastRegistryRefresh = std::chrono::steady_clock::now();

	for (;;) {
		if (g_stopRequested.load()) {
			break;
		}
		if (std::chrono::steady_clock::now() - lastRegistryRefresh >= std::chrono::seconds(2)) {
			RefreshCurrentInstanceRegistry();
			lastRegistryRefresh = std::chrono::steady_clock::now();
		}

		SOCKET listenSock = INVALID_SOCKET;
		{
			std::lock_guard<std::mutex> lock(g_stateMutex);
			listenSock = g_listenSocket;
		}
		if (listenSock == INVALID_SOCKET) {
			break;
		}

		fd_set readSet;
		FD_ZERO(&readSet);
		FD_SET(listenSock, &readSet);
		timeval timeout = {};
		timeout.tv_sec = 0;
		timeout.tv_usec = 250000;
		const int ready = select(0, &readSet, nullptr, nullptr, &timeout);
		if (ready == SOCKET_ERROR) {
			const int error = WSAGetLastError();
			if (!g_stopRequested.load()) {
				LogMcp(std::format("select failed, error={}", error));
			}
			break;
		}
		if (ready == 0) {
			continue;
		}

		sockaddr_in clientAddr = {};
		int clientLen = sizeof(clientAddr);
		SOCKET clientSock = accept(listenSock, reinterpret_cast<sockaddr*>(&clientAddr), &clientLen);
		if (clientSock == INVALID_SOCKET) {
			if (!g_stopRequested.load()) {
				LogMcp(std::format("accept failed, error={}", WSAGetLastError()));
			}
			continue;
		}

		if (!TryQueueClient(clientSock)) {
			SendHttpResponse(
				clientSock,
				503,
				"Service Unavailable",
				"application/json; charset=utf-8",
				R"({"ok":false,"error":"server busy"})");
			CloseSocketSafe(clientSock);
		}
	}

	StopClientWorkersAndCloseQueue();
	for (std::thread& worker : clientWorkers) {
		if (worker.joinable()) {
			worker.join();
		}
	}
	{
		std::lock_guard<std::mutex> lock(g_stateMutex);
		CloseSocketSafe(g_listenSocket);
	}
	g_running.store(false);
	g_boundPort.store(0);
	RemoveCurrentInstanceRegistry();
	WSACleanup();
	LogMcp("stopped");
}

} // namespace

namespace LocalMcpServer {

void Initialize()
{
	std::lock_guard<std::mutex> lock(g_stateMutex);
	if (g_serverThread.joinable()) {
		return;
	}
	g_stopRequested.store(false);
	g_running.store(false);
	g_registryRefreshFailed.store(false);
	g_boundPort.store(0);
	g_instanceId = GenerateInstanceId();
	g_sourceFilePathHint.clear();
	g_pageNameHint.clear();
	g_pageTypeHint.clear();
	g_externalMcpSessions.clear();

	// FneInit 通常已打开主日志；直接初始化 MCP 时再兜底打开。
	Logger::Instance().OpenIfNeeded(GetAutoLinkerLogFilePath("autolinker.log").string());

	g_serverThread = std::thread(ServerThreadMain);
}

void Shutdown()
{
	std::thread worker;
	{
		std::lock_guard<std::mutex> lock(g_stateMutex);
		if (!g_serverThread.joinable()) {
			return;
		}
		g_stopRequested.store(true);
		CloseSocketSafe(g_listenSocket);
		worker = std::move(g_serverThread);
	}
	StopClientWorkersAndCloseQueue();
	if (worker.joinable()) {
		worker.join();
	}
}

bool IsRunning()
{
	return g_running.load();
}

int GetBoundPort()
{
	return g_boundPort.load();
}

std::string GetInstanceId()
{
	std::lock_guard<std::mutex> lock(g_stateMutex);
	return g_instanceId;
}

std::string GetEndpoint()
{
	return BuildEndpointForPort(g_boundPort.load());
}

std::string BuildWorkspaceRefreshGateSelfTestJson()
{
	const std::string sessionId = "self-test-refresh-gate";
	SetExternalWorkspaceRefreshed(sessionId, false);
	nlohmann::json result;
	std::string error;
	const bool handled = TryBuildToolCallResult(
		{
			{"name", "read_files"},
			{"arguments", {{"file_paths", nlohmann::json::array({"src/Test.txt"})}}}
		},
		sessionId,
		result,
		error);
	{
		std::lock_guard<std::mutex> lock(g_stateMutex);
		g_externalMcpSessions.erase(sessionId);
	}
	const bool blocked = handled &&
		result.value("isError", false) &&
		result.contains("structuredContent") &&
		result["structuredContent"].is_object() &&
		result["structuredContent"].value("error", std::string()) == "workspace_refresh_required";
	nlohmann::json hiddenResult;
	std::string hiddenError;
	const bool hiddenHandled = TryBuildToolCallResult(
		{
			{"name", "update_plan"},
			{"arguments", {{"plan", nlohmann::json::array()}}}
		},
		sessionId,
		hiddenResult,
		hiddenError);
	const bool hiddenToolBlocked = !hiddenHandled &&
		hiddenError.find("unknown or unavailable external tool") != std::string::npos;
	nlohmann::json missingHashResult;
	std::string missingHashError;
	const bool missingHashHandled = TryBuildToolCallResult(
		{
			{"name", "edit_file"},
			{"arguments", {
				{"file_path", "src/Test.txt"},
				{"old_text", "a"},
				{"new_text", "b"}
			}}
		},
		sessionId,
		missingHashResult,
		missingHashError);
	const bool missingHashBlocked = !missingHashHandled &&
		missingHashError.find("expected_base_hash") != std::string::npos;
	return nlohmann::json({
		{"name", "external_mcp_workspace_refresh_gate"},
		{"ok", blocked && hiddenToolBlocked && missingHashBlocked},
		{"handled", handled},
		{"error", error},
		{"result", result},
		{"hidden_tool_blocked", hiddenToolBlocked},
		{"hidden_tool_error", hiddenError},
		{"missing_hash_blocked", missingHashBlocked},
		{"missing_hash_error", missingHashError}
	}).dump();
}

void UpdateInstanceHints(const std::string& sourceFilePath, const std::string& pageName, const std::string& pageType)
{
	{
		std::lock_guard<std::mutex> lock(g_stateMutex);
		if (g_sourceFilePathHint != sourceFilePath) {
			ResetExternalWorkspaceRefreshStatesLocked();
		}
		g_sourceFilePathHint = sourceFilePath;
		g_pageNameHint = pageName;
		g_pageTypeHint = pageType;
	}
	if (g_running.load()) {
		RefreshCurrentInstanceRegistry();
	}
}

} // namespace LocalMcpServer
