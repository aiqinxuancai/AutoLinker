#include "LocalMcpServer.h"

#include <winsock2.h>
#include <ws2tcpip.h>

#include <algorithm>
#include <atomic>
#include <cctype>
#include <chrono>
#include <cstdlib>
#include <format>
#include <mutex>
#include <string>
#include <thread>
#include <unordered_map>
#include <vector>

#include "..\\thirdparty\\json.hpp"

#include "AIChatFeature.h"
#include "AIService.h"
#include "Global.h"
#include "IDEFacade.h"

#pragma comment(lib, "Ws2_32.lib")

namespace {

constexpr const char* kServerName = "AutoLinker Local MCP";
constexpr const char* kServerVersion = "0.0.0";
constexpr const char* kBindHost = "127.0.0.1";
constexpr int kBasePort = 19207;
constexpr int kMaxPortAttempts = 16;

std::atomic_bool g_stopRequested = false;
std::atomic_bool g_running = false;
std::atomic_int g_boundPort = 0;
std::mutex g_stateMutex;
std::thread g_serverThread;
SOCKET g_listenSocket = INVALID_SOCKET;

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
	OutputStringToELog("[LocalMCP] " + message);
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

void LogMcpCallLine(const std::string& message)
{
	OutputStringToELog(ConvertUtf8ToGbkText("[MCP] " + message));
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

std::string FormatMcpLogJson(const nlohmann::json& value)
{
	return TruncateMcpLogText(SanitizeSingleLineText(value.dump()), 180);
}

std::string FormatMcpLogText(const std::string& value)
{
	return TruncateMcpLogText(SanitizeSingleLineText(value), 180);
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

struct McpLogContext {
	bool enabled = false;
	std::string responseName;
	std::string requestDisplay;
	std::string responseDisplay;
};

McpLogContext BuildMcpLogContextForPayload(const nlohmann::json& payload)
{
	McpLogContext ctx{};
	ctx.enabled = true;

	if (!payload.is_object()) {
		ctx.responseName = "invalid_request";
		ctx.requestDisplay = "invalid_request(" + FormatMcpLogJson(payload) + ")";
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
		ctx.requestDisplay = toolName + "(" + BuildJsonValueCallSuffix(arguments) + ")";
		return ctx;
	}

	ctx.responseName = method;
	ctx.requestDisplay = method + "(" + BuildJsonValueCallSuffix(params) + ")";
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
		ctx.requestDisplay = "invalid_json(" + FormatMcpLogText(body) + ")";
		return ctx;
	}
}

void LogMcpRequest(const McpLogContext& ctx)
{
	if (!ctx.enabled) {
		return;
	}
	LogMcpCallLine(">> " + ctx.requestDisplay);
}

void LogMcpResponse(const McpLogContext& ctx, double elapsedMs)
{
	if (!ctx.enabled) {
		return;
	}

	const std::string responseText = ctx.responseDisplay.empty() ? "null" : ctx.responseDisplay;
	LogMcpCallLine(std::format(
		"<< {} ({:.1f}ms) {}",
		ctx.responseName.empty() ? "unknown" : ctx.responseName,
		elapsedMs,
		responseText));
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
	const std::string& body)
{
	return std::format(
		"HTTP/1.1 {} {}\r\n"
		"Content-Type: {}\r\n"
		"Content-Length: {}\r\n"
		"Connection: close\r\n"
		"Cache-Control: no-store\r\n"
		"Access-Control-Allow-Origin: *\r\n"
		"Access-Control-Allow-Headers: content-type, mcp-session-id\r\n"
		"Access-Control-Allow-Methods: GET, POST, OPTIONS\r\n"
		"\r\n{}",
		statusCode,
		statusText,
		contentType,
		body.size(),
		body);
}

void SendHttpResponse(
	SOCKET sock,
	int statusCode,
	const char* statusText,
	const char* contentType,
	const std::string& body)
{
	const std::string response = BuildHttpResponse(statusCode, statusText, contentType, body);
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

nlohmann::json BuildInitializeResult()
{
	return {
		{"protocolVersion", "2024-11-05"},
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
		outResult = {
			{"tools", nlohmann::json::parse(AIService::BuildPublicToolCatalogJson())}
		};
		return true;
	}
	catch (const std::exception& ex) {
		outError = std::string("build tools list failed: ") + ex.what();
		return false;
	}
}

bool TryBuildToolCallResult(const nlohmann::json& params, nlohmann::json& outResult, std::string& outError)
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

	std::string resultJsonUtf8;
	bool toolOk = false;
	if (!AIChatFeature::ExecutePublicTool(toolName, arguments.dump(), resultJsonUtf8, toolOk)) {
		outError = "tool execution transport failed";
		return false;
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

	outResult = {
		{"content", nlohmann::json::array({
			{
				{"type", "text"},
				{"text", resultJsonUtf8}
			}
		})},
		{"isError", isError}
	};
	if (structuredOk) {
		outResult["structuredContent"] = structured;
	}
	return true;
}

bool TryHandleJsonRpc(const HttpRequest& request, int& outStatusCode, std::string& outBody, McpLogContext* outLogContext)
{
	outStatusCode = 200;
	outBody.clear();
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
			outLogContext->requestDisplay = "invalid_json(" + FormatMcpLogText(request.body) + ")";
			outLogContext->responseDisplay = FormatMcpLogJson({
				{"code", -32700},
				{"message", std::string("parse error: ") + ex.what()}
			});
		}
		outBody = BuildJsonRpcError(nullptr, -32700, std::string("parse error: ") + ex.what()).dump();
		return true;
	}

	if (outLogContext != nullptr) {
		*outLogContext = BuildMcpLogContextForPayload(payload);
	}

	if (!payload.is_object()) {
		if (outLogContext != nullptr) {
			outLogContext->responseDisplay = FormatMcpLogJson({
				{"code", -32600},
				{"message", "request must be a JSON object"}
			});
		}
		outBody = BuildJsonRpcError(nullptr, -32600, "request must be a JSON object").dump();
		return true;
	}

	const nlohmann::json id = payload.contains("id") ? payload["id"] : nlohmann::json(nullptr);
	if (!payload.contains("method") || !payload["method"].is_string()) {
		if (outLogContext != nullptr) {
			outLogContext->responseDisplay = FormatMcpLogJson({
				{"code", -32600},
				{"message", "method is required"}
			});
		}
		outBody = BuildJsonRpcError(id, -32600, "method is required").dump();
		return true;
	}

	const std::string method = payload["method"].get<std::string>();
	const bool hasId = payload.contains("id");
	const nlohmann::json params = payload.contains("params") ? payload["params"] : nlohmann::json::object();

	if (method == "notifications/initialized") {
		outStatusCode = 202;
		outBody.clear();
		if (outLogContext != nullptr) {
			outLogContext->responseDisplay = "null";
		}
		return true;
	}

	if (method == "ping") {
		const nlohmann::json result = nlohmann::json::object();
		if (outLogContext != nullptr) {
			outLogContext->responseDisplay = FormatMcpLogJson(result);
		}
		outBody = BuildJsonRpcResult(id, result).dump();
		return true;
	}

	if (method == "initialize") {
		const nlohmann::json result = BuildInitializeResult();
		if (outLogContext != nullptr) {
			outLogContext->responseDisplay = FormatMcpLogJson(result);
		}
		outBody = BuildJsonRpcResult(id, result).dump();
		return true;
	}

	if (method == "tools/list") {
		nlohmann::json result;
		std::string error;
		if (!TryBuildToolListResult(result, error)) {
			if (outLogContext != nullptr) {
				outLogContext->responseDisplay = FormatMcpLogJson({
					{"code", -32603},
					{"message", error}
				});
			}
			outBody = BuildJsonRpcError(id, -32603, error).dump();
			return true;
		}
		if (outLogContext != nullptr) {
			outLogContext->responseDisplay = FormatMcpLogJson(result);
		}
		outBody = BuildJsonRpcResult(id, result).dump();
		return true;
	}

	if (method == "tools/call") {
		nlohmann::json result;
		std::string error;
		if (!TryBuildToolCallResult(params, result, error)) {
			if (outLogContext != nullptr) {
				outLogContext->responseDisplay = FormatMcpLogJson({
					{"code", -32602},
					{"message", error}
				});
			}
			outBody = BuildJsonRpcError(id, -32602, error).dump();
			return true;
		}
		if (outLogContext != nullptr) {
			outLogContext->responseDisplay = FormatMcpLogJson(result);
		}
		outBody = BuildJsonRpcResult(id, result).dump();
		return true;
	}

	if (!hasId) {
		outStatusCode = 202;
		outBody.clear();
		if (outLogContext != nullptr) {
			outLogContext->responseDisplay = "null";
		}
		return true;
	}

	if (outLogContext != nullptr) {
		outLogContext->responseDisplay = FormatMcpLogJson({
			{"code", -32601},
			{"message", "method not found"}
		});
	}
	outBody = BuildJsonRpcError(id, -32601, "method not found").dump();
	return true;
}

void HandleClient(SOCKET clientSock)
{
	DWORD timeoutMs = 5000;
	setsockopt(clientSock, SOL_SOCKET, SO_RCVTIMEO, reinterpret_cast<const char*>(&timeoutMs), sizeof(timeoutMs));
	setsockopt(clientSock, SOL_SOCKET, SO_SNDTIMEO, reinterpret_cast<const char*>(&timeoutMs), sizeof(timeoutMs));

	HttpRequest request;
	if (!ReadHttpRequest(clientSock, request)) {
		SendHttpResponse(clientSock, 400, "Bad Request", "application/json; charset=utf-8", R"({"ok":false,"error":"invalid http request"})");
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
			{"port", g_boundPort.load()},
			{"mcp_endpoint", std::format("http://{}:{}/mcp", kBindHost, g_boundPort.load())}
		};
		SendHttpResponse(clientSock, 200, "OK", "application/json; charset=utf-8", health.dump());
		return;
	}

	if (request.method != "POST") {
		SendHttpResponse(clientSock, 405, "Method Not Allowed", "application/json; charset=utf-8", R"({"ok":false,"error":"method not allowed"})");
		return;
	}

	if (request.path != "/" && request.path != "/mcp") {
		SendHttpResponse(clientSock, 404, "Not Found", "application/json; charset=utf-8", R"({"ok":false,"error":"not found"})");
		return;
	}

	McpLogContext logContext = BuildMcpLogContextForRequestBody(request.body);
	LogMcpRequest(logContext);
	const auto startTime = std::chrono::steady_clock::now();

	int statusCode = 200;
	std::string responseBody;
	McpLogContext handledLogContext;
	if (!TryHandleJsonRpc(request, statusCode, responseBody, &handledLogContext)) {
		const double elapsedMs = std::chrono::duration<double, std::milli>(
			std::chrono::steady_clock::now() - startTime).count();
		logContext.responseName = handledLogContext.responseName.empty() ? logContext.responseName : handledLogContext.responseName;
		logContext.responseDisplay = R"({"ok":false,"error":"internal server error"})";
		LogMcpResponse(logContext, elapsedMs);
		SendHttpResponse(clientSock, 500, "Internal Server Error", "application/json; charset=utf-8", R"({"ok":false,"error":"internal server error"})");
		return;
	}
	if (handledLogContext.enabled) {
		if (!handledLogContext.responseName.empty()) {
			logContext.responseName = handledLogContext.responseName;
		}
		if (!handledLogContext.requestDisplay.empty()) {
			logContext.requestDisplay = handledLogContext.requestDisplay;
		}
		logContext.responseDisplay = handledLogContext.responseDisplay;
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
	SendHttpResponse(clientSock, statusCode, statusText, "application/json; charset=utf-8", responseBody);
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
	g_running.store(true);
	LogMcp(std::format("listening on http://{}:{}/mcp", kBindHost, boundPort));

	for (;;) {
		if (g_stopRequested.load()) {
			break;
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

		HandleClient(clientSock);
		CloseSocketSafe(clientSock);
	}

	{
		std::lock_guard<std::mutex> lock(g_stateMutex);
		CloseSocketSafe(g_listenSocket);
	}
	g_running.store(false);
	g_boundPort.store(0);
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
	g_boundPort.store(0);
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

} // namespace LocalMcpServer
