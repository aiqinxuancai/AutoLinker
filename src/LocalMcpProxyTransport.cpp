#include "LocalMcpProxyTransport.h"

#include <winsock2.h>
#include <ws2tcpip.h>

#include <algorithm>
#include <cctype>
#include <cstdlib>
#include <format>
#include <string>
#include <thread>
#include <unordered_map>

#include "..\\thirdparty\\json.hpp"

#pragma comment(lib, "Ws2_32.lib")

namespace {

constexpr const char* kLoopbackHost = "127.0.0.1";
constexpr std::size_t kMaxHeaderBytes = 64 * 1024;
constexpr std::size_t kMaxBodyBytes = 4 * 1024 * 1024;
constexpr const char* kProxiedHeaderName = "X-AutoLinker-MCP-Proxied";

class WinsockScope {
public:
	WinsockScope()
	{
		WSADATA data = {};
		m_started = WSAStartup(MAKEWORD(2, 2), &data) == 0;
	}

	~WinsockScope()
	{
		if (m_started) {
			WSACleanup();
		}
	}

	bool Started() const { return m_started; }

private:
	bool m_started = false;
};

class SocketScope {
public:
	explicit SocketScope(SOCKET socketValue) : m_socket(socketValue) {}
	~SocketScope()
	{
		if (m_socket != INVALID_SOCKET) {
			shutdown(m_socket, SD_BOTH);
			closesocket(m_socket);
		}
	}
	SOCKET Get() const { return m_socket; }

private:
	SOCKET m_socket = INVALID_SOCKET;
};

std::string TrimAsciiCopy(const std::string& text)
{
	std::size_t begin = 0;
	std::size_t end = text.size();
	while (begin < end && std::isspace(static_cast<unsigned char>(text[begin])) != 0) {
		++begin;
	}
	while (end > begin && std::isspace(static_cast<unsigned char>(text[end - 1])) != 0) {
		--end;
	}
	return text.substr(begin, end - begin);
}

std::string ToLowerAsciiCopy(std::string text)
{
	std::transform(text.begin(), text.end(), text.begin(), [](unsigned char ch) {
		return static_cast<char>(std::tolower(ch));
	});
	return text;
}

bool ConnectWithTimeout(SOCKET socketValue, int port, int timeoutMs, std::string& outError)
{
	u_long nonBlocking = 1;
	if (ioctlsocket(socketValue, FIONBIO, &nonBlocking) == SOCKET_ERROR) {
		outError = "enable nonblocking socket failed, error=" + std::to_string(WSAGetLastError());
		return false;
	}

	sockaddr_in address = {};
	address.sin_family = AF_INET;
	address.sin_port = htons(static_cast<u_short>(port));
	if (inet_pton(AF_INET, kLoopbackHost, &address.sin_addr) != 1) {
		outError = "parse loopback address failed";
		return false;
	}

	const int connectResult = connect(
		socketValue,
		reinterpret_cast<const sockaddr*>(&address),
		sizeof(address));
	if (connectResult == SOCKET_ERROR) {
		const int connectError = WSAGetLastError();
		if (connectError != WSAEWOULDBLOCK && connectError != WSAEINPROGRESS) {
			outError = "connect loopback instance failed, error=" + std::to_string(connectError);
			return false;
		}

		fd_set writeSet;
		FD_ZERO(&writeSet);
		FD_SET(socketValue, &writeSet);
		timeval timeout = {};
		timeout.tv_sec = timeoutMs / 1000;
		timeout.tv_usec = (timeoutMs % 1000) * 1000;
		const int selected = select(0, nullptr, &writeSet, nullptr, &timeout);
		if (selected <= 0) {
			outError = selected == 0
				? "connect loopback instance timed out"
				: "wait for loopback connection failed, error=" + std::to_string(WSAGetLastError());
			return false;
		}

		int socketError = 0;
		int socketErrorLength = sizeof(socketError);
		if (getsockopt(
			socketValue,
			SOL_SOCKET,
			SO_ERROR,
			reinterpret_cast<char*>(&socketError),
			&socketErrorLength) == SOCKET_ERROR ||
			socketError != 0) {
			outError = "connect loopback instance failed, error=" + std::to_string(
				socketError != 0 ? socketError : WSAGetLastError());
			return false;
		}
	}

	nonBlocking = 0;
	if (ioctlsocket(socketValue, FIONBIO, &nonBlocking) == SOCKET_ERROR) {
		outError = "restore blocking socket failed, error=" + std::to_string(WSAGetLastError());
		return false;
	}
	return true;
}

bool SendAll(SOCKET socketValue, const std::string& data, std::string& outError)
{
	std::size_t sentTotal = 0;
	while (sentTotal < data.size()) {
		const int sent = send(
			socketValue,
			data.data() + sentTotal,
			static_cast<int>((std::min)(data.size() - sentTotal, static_cast<std::size_t>(INT_MAX))),
			0);
		if (sent <= 0) {
			outError = "send loopback request failed, error=" + std::to_string(WSAGetLastError());
			return false;
		}
		sentTotal += static_cast<std::size_t>(sent);
	}
	return true;
}

bool ReadHttpResponse(
	SOCKET socketValue,
	LocalMcpProxyTransport::HttpResponse& outResponse,
	std::string& outError)
{
	std::string raw;
	std::size_t headerEnd = std::string::npos;
	while ((headerEnd = raw.find("\r\n\r\n")) == std::string::npos) {
		char buffer[4096];
		const int received = recv(socketValue, buffer, static_cast<int>(sizeof(buffer)), 0);
		if (received <= 0) {
			outError = "read loopback response header failed, error=" + std::to_string(WSAGetLastError());
			return false;
		}
		raw.append(buffer, static_cast<std::size_t>(received));
		if (raw.size() > kMaxHeaderBytes) {
			outError = "loopback response header is too large";
			return false;
		}
	}

	const std::string headerText = raw.substr(0, headerEnd);
	std::string body = raw.substr(headerEnd + 4);
	const std::size_t statusLineEnd = headerText.find("\r\n");
	const std::string statusLine = statusLineEnd == std::string::npos
		? headerText
		: headerText.substr(0, statusLineEnd);
	const std::size_t firstSpace = statusLine.find(' ');
	const std::size_t secondSpace = firstSpace == std::string::npos
		? std::string::npos
		: statusLine.find(' ', firstSpace + 1);
	if (firstSpace == std::string::npos || secondSpace == std::string::npos) {
		outError = "invalid loopback HTTP status line";
		return false;
	}
	outResponse.statusCode = std::atoi(statusLine.substr(firstSpace + 1, secondSpace - firstSpace - 1).c_str());
	outResponse.reason = TrimAsciiCopy(statusLine.substr(secondSpace + 1));
	if (outResponse.statusCode <= 0) {
		outError = "invalid loopback HTTP status code";
		return false;
	}

	std::unordered_map<std::string, std::string> headers;
	std::size_t lineBegin = statusLineEnd == std::string::npos ? headerText.size() : statusLineEnd + 2;
	while (lineBegin < headerText.size()) {
		const std::size_t lineEnd = headerText.find("\r\n", lineBegin);
		const std::string line = lineEnd == std::string::npos
			? headerText.substr(lineBegin)
			: headerText.substr(lineBegin, lineEnd - lineBegin);
		lineBegin = lineEnd == std::string::npos ? headerText.size() : lineEnd + 2;
		const std::size_t colon = line.find(':');
		if (colon != std::string::npos) {
			headers[ToLowerAsciiCopy(TrimAsciiCopy(line.substr(0, colon)))] =
				TrimAsciiCopy(line.substr(colon + 1));
		}
	}

	const auto lengthIt = headers.find("content-length");
	if (lengthIt == headers.end()) {
		outError = "loopback response is missing Content-Length";
		return false;
	}
	const unsigned long long contentLength = std::strtoull(lengthIt->second.c_str(), nullptr, 10);
	if (contentLength > kMaxBodyBytes) {
		outError = "loopback response body is too large";
		return false;
	}
	while (body.size() < contentLength) {
		char buffer[4096];
		const int received = recv(socketValue, buffer, static_cast<int>(sizeof(buffer)), 0);
		if (received <= 0) {
			outError = "read loopback response body failed, error=" + std::to_string(WSAGetLastError());
			return false;
		}
		body.append(buffer, static_cast<std::size_t>(received));
		if (body.size() > kMaxBodyBytes) {
			outError = "loopback response body is too large";
			return false;
		}
	}
	outResponse.body = body.substr(0, static_cast<std::size_t>(contentLength));
	return true;
}

bool SendLoopbackRequest(
	int port,
	const std::string& method,
	const std::string& body,
	const std::string& sessionId,
	bool proxied,
	LocalMcpProxyTransport::HttpResponse& outResponse,
	std::string& outError,
	int timeoutMs)
{
	outResponse = {};
	outError.clear();
	if (port <= 0 || port > 65535) {
		outError = "loopback instance port is invalid";
		return false;
	}
	if (timeoutMs <= 0) {
		outError = "loopback timeout is invalid";
		return false;
	}

	WinsockScope winsock;
	if (!winsock.Started()) {
		outError = "WSAStartup failed";
		return false;
	}

	SocketScope socketScope(socket(AF_INET, SOCK_STREAM, IPPROTO_TCP));
	if (socketScope.Get() == INVALID_SOCKET) {
		outError = "create loopback socket failed, error=" + std::to_string(WSAGetLastError());
		return false;
	}
	DWORD socketTimeout = static_cast<DWORD>(timeoutMs);
	setsockopt(socketScope.Get(), SOL_SOCKET, SO_RCVTIMEO, reinterpret_cast<const char*>(&socketTimeout), sizeof(socketTimeout));
	setsockopt(socketScope.Get(), SOL_SOCKET, SO_SNDTIMEO, reinterpret_cast<const char*>(&socketTimeout), sizeof(socketTimeout));
	if (!ConnectWithTimeout(socketScope.Get(), port, timeoutMs, outError)) {
		return false;
	}

	std::string headers;
	if (!sessionId.empty()) {
		headers += "Mcp-Session-Id: " + sessionId + "\r\n";
	}
	if (proxied) {
		headers += std::string(kProxiedHeaderName) + ": 1\r\n";
	}
	const std::string request = std::format(
		"{} /mcp HTTP/1.1\r\n"
		"Host: {}:{}\r\n"
		"Content-Type: application/json\r\n"
		"Content-Length: {}\r\n"
		"Connection: close\r\n"
		"{}"
		"\r\n{}",
		method,
		kLoopbackHost,
		port,
		body.size(),
		headers,
		body);
	return SendAll(socketScope.Get(), request, outError) &&
		ReadHttpResponse(socketScope.Get(), outResponse, outError);
}

bool ReadMockRequest(SOCKET socketValue, std::string& outRaw)
{
	outRaw.clear();
	std::size_t headerEnd = std::string::npos;
	std::size_t contentLength = 0;
	for (;;) {
		char buffer[2048];
		const int received = recv(socketValue, buffer, static_cast<int>(sizeof(buffer)), 0);
		if (received <= 0) {
			return false;
		}
		outRaw.append(buffer, static_cast<std::size_t>(received));
		if (outRaw.size() > kMaxHeaderBytes + 4096) {
			return false;
		}
		if (headerEnd == std::string::npos) {
			headerEnd = outRaw.find("\r\n\r\n");
			if (headerEnd != std::string::npos) {
				const std::string headerText = ToLowerAsciiCopy(outRaw.substr(0, headerEnd));
				const std::size_t lengthPos = headerText.find("content-length:");
				if (lengthPos != std::string::npos) {
					const std::size_t valueBegin = lengthPos + std::string("content-length:").size();
					const std::size_t valueEnd = headerText.find("\r\n", valueBegin);
					contentLength = static_cast<std::size_t>(std::strtoull(
						TrimAsciiCopy(headerText.substr(valueBegin, valueEnd - valueBegin)).c_str(),
						nullptr,
						10));
				}
			}
		}
		if (headerEnd != std::string::npos && outRaw.size() >= headerEnd + 4 + contentLength) {
			return true;
		}
	}
}

void SendMockResponse(SOCKET socketValue, const std::string& body)
{
	const std::string response = std::format(
		"HTTP/1.1 200 OK\r\n"
		"Content-Type: application/json\r\n"
		"Content-Length: {}\r\n"
		"Connection: close\r\n"
		"\r\n{}",
		body.size(),
		body);
	std::string ignored;
	SendAll(socketValue, response, ignored);
}

} // namespace

namespace LocalMcpProxyTransport {

bool PostJsonRpc(
	int port,
	const std::string& requestBody,
	const std::string& sessionId,
	HttpResponse& outResponse,
	std::string& outError,
	int timeoutMs)
{
	return SendLoopbackRequest(
		port,
		"POST",
		requestBody,
		sessionId,
		true,
		outResponse,
		outError,
		timeoutMs);
}

bool ProbeInstance(
	int port,
	const std::string& expectedInstanceId,
	std::string& outError,
	int timeoutMs)
{
	HttpResponse response;
	if (!SendLoopbackRequest(port, "GET", std::string(), std::string(), false, response, outError, timeoutMs)) {
		return false;
	}
	if (response.statusCode != 200) {
		outError = std::format("instance health request returned HTTP {} {}", response.statusCode, response.reason);
		return false;
	}
	const nlohmann::json health = nlohmann::json::parse(response.body, nullptr, false);
	if (health.is_discarded() || !health.is_object() || !health.value("ok", false)) {
		outError = "instance health response is invalid";
		return false;
	}
	if (!expectedInstanceId.empty() && health.value("instance_id", std::string()) != expectedInstanceId) {
		outError = "instance health identity does not match registry";
		return false;
	}
	return true;
}

std::string BuildSelfTestReportJson()
{
	nlohmann::json report = {
		{"name", "local-mcp-loopback-transport"},
		{"ok", false}
	};
	WinsockScope winsock;
	if (!winsock.Started()) {
		report["error"] = "WSAStartup failed";
		return report.dump();
	}

	SOCKET listenSocket = socket(AF_INET, SOCK_STREAM, IPPROTO_TCP);
	if (listenSocket == INVALID_SOCKET) {
		report["error"] = "create mock socket failed";
		return report.dump();
	}
	SocketScope listenScope(listenSocket);
	sockaddr_in address = {};
	address.sin_family = AF_INET;
	address.sin_port = 0;
	inet_pton(AF_INET, kLoopbackHost, &address.sin_addr);
	if (bind(listenSocket, reinterpret_cast<const sockaddr*>(&address), sizeof(address)) == SOCKET_ERROR ||
		listen(listenSocket, 2) == SOCKET_ERROR) {
		report["error"] = "start mock server failed";
		return report.dump();
	}
	int addressLength = sizeof(address);
	if (getsockname(listenSocket, reinterpret_cast<sockaddr*>(&address), &addressLength) == SOCKET_ERROR) {
		report["error"] = "resolve mock server port failed";
		return report.dump();
	}
	const int port = ntohs(address.sin_port);
	std::string getRequest;
	std::string postRequest;
	std::thread mockServer([&]() {
		for (int requestIndex = 0; requestIndex < 2; ++requestIndex) {
			SOCKET accepted = accept(listenSocket, nullptr, nullptr);
			if (accepted == INVALID_SOCKET) {
				return;
			}
			SocketScope acceptedScope(accepted);
			DWORD timeoutMs = 5000;
			setsockopt(accepted, SOL_SOCKET, SO_RCVTIMEO, reinterpret_cast<const char*>(&timeoutMs), sizeof(timeoutMs));
			std::string raw;
			if (!ReadMockRequest(accepted, raw)) {
				return;
			}
			if (raw.starts_with("GET ")) {
				getRequest = raw;
				SendMockResponse(accepted, R"({"ok":true,"instance_id":"mock-instance"})");
			}
			else {
				postRequest = raw;
				SendMockResponse(accepted, R"({"jsonrpc":"2.0","id":7,"result":{"ok":true}})");
			}
		}
	});

	std::string probeError;
	const bool probeOk = ProbeInstance(port, "mock-instance", probeError, 2000);
	HttpResponse response;
	std::string postError;
	const bool postOk = PostJsonRpc(
		port,
		R"({"jsonrpc":"2.0","id":7,"method":"ping"})",
		"session-test",
		response,
		postError,
		2000);
	shutdown(listenSocket, SD_BOTH);
	if (mockServer.joinable()) {
		mockServer.join();
	}
	const std::string loweredPost = ToLowerAsciiCopy(postRequest);
	const bool headersOk =
		loweredPost.find("mcp-session-id: session-test") != std::string::npos &&
		loweredPost.find("x-autolinker-mcp-proxied: 1") != std::string::npos;
	const nlohmann::json responseJson = nlohmann::json::parse(response.body, nullptr, false);
	const bool responseOk = postOk && response.statusCode == 200 &&
		!responseJson.is_discarded() && responseJson.value("jsonrpc", std::string()) == "2.0";
	report["ok"] = probeOk && responseOk && headersOk && !getRequest.empty();
	report["probe_ok"] = probeOk;
	report["post_ok"] = responseOk;
	report["headers_ok"] = headersOk;
	if (!probeError.empty()) {
		report["probe_error"] = probeError;
	}
	if (!postError.empty()) {
		report["post_error"] = postError;
	}
	return report.dump();
}

} // namespace LocalMcpProxyTransport
