#pragma once

#include <string>

namespace LocalMcpProxyTransport {

// 本机 MCP 回环请求结果
struct HttpResponse {
	int statusCode = 0;
	std::string reason;
	std::string body;
};

// 向已登记的本机 MCP 实例转发 JSON-RPC 请求
bool PostJsonRpc(
	int port,
	const std::string& requestBody,
	const std::string& sessionId,
	HttpResponse& outResponse,
	std::string& outError,
	int timeoutMs = 30000);

// 验证本机 MCP 实例端口及实例标识
bool ProbeInstance(
	int port,
	const std::string& expectedInstanceId,
	std::string& outError,
	int timeoutMs = 500);

// 构建回环转发协议的内部自测报告
std::string BuildSelfTestReportJson();

} // namespace LocalMcpProxyTransport
