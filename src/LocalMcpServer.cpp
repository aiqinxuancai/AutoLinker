#include "LocalMcpServer.h"

#include <winsock2.h>
#include <ws2tcpip.h>

#include <algorithm>
#include <atomic>
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

bool TryHandleJsonRpc(const HttpRequest& request, int& outStatusCode, std::string& outBody)
{
	outStatusCode = 200;
	outBody.clear();

	nlohmann::json payload;
	try {
		payload = request.body.empty() ? nlohmann::json::object() : nlohmann::json::parse(request.body);
	}
	catch (const std::exception& ex) {
		outBody = BuildJsonRpcError(nullptr, -32700, std::string("parse error: ") + ex.what()).dump();
		return true;
	}

	if (!payload.is_object()) {
		outBody = BuildJsonRpcError(nullptr, -32600, "request must be a JSON object").dump();
		return true;
	}

	const nlohmann::json id = payload.contains("id") ? payload["id"] : nlohmann::json(nullptr);
	if (!payload.contains("method") || !payload["method"].is_string()) {
		outBody = BuildJsonRpcError(id, -32600, "method is required").dump();
		return true;
	}

	const std::string method = payload["method"].get<std::string>();
	const bool hasId = payload.contains("id");
	const nlohmann::json params = payload.contains("params") ? payload["params"] : nlohmann::json::object();

	if (method == "notifications/initialized") {
		outStatusCode = 202;
		outBody.clear();
		return true;
	}

	if (method == "ping") {
		outBody = BuildJsonRpcResult(id, nlohmann::json::object()).dump();
		return true;
	}

	if (method == "initialize") {
		outBody = BuildJsonRpcResult(id, BuildInitializeResult()).dump();
		return true;
	}

	if (method == "tools/list") {
		nlohmann::json result;
		std::string error;
		if (!TryBuildToolListResult(result, error)) {
			outBody = BuildJsonRpcError(id, -32603, error).dump();
			return true;
		}
		outBody = BuildJsonRpcResult(id, result).dump();
		return true;
	}

	if (method == "tools/call") {
		nlohmann::json result;
		std::string error;
		if (!TryBuildToolCallResult(params, result, error)) {
			outBody = BuildJsonRpcError(id, -32602, error).dump();
			return true;
		}
		outBody = BuildJsonRpcResult(id, result).dump();
		return true;
	}

	if (!hasId) {
		outStatusCode = 202;
		outBody.clear();
		return true;
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

	int statusCode = 200;
	std::string responseBody;
	if (!TryHandleJsonRpc(request, statusCode, responseBody)) {
		SendHttpResponse(clientSock, 500, "Internal Server Error", "application/json; charset=utf-8", R"({"ok":false,"error":"internal server error"})");
		return;
	}

	const char* statusText = "OK";
	if (statusCode == 202) {
		statusText = "Accepted";
	}
	else if (statusCode == 204) {
		statusText = "No Content";
	}
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
