#pragma once

#include <string>
#include "ToolExecutionTimeouts.h"

namespace LocalMcpProxyTransport {

// 本机 MCP 回环请求结果
struct HttpResponse {
	int statusCode = 0;
	std::string reason;
	std::string body;
	bool requestSent = false; // 请求已完整发送，之后失败不能推断工具未执行。
	bool responseTimedOut = false; // 等待响应头或响应正文超时。
};

// 向已登记的本机 MCP 实例转发 JSON-RPC 请求
bool PostJsonRpc(
	int port,
	const std::string& requestBody,
	const std::string& sessionId,
	HttpResponse& outResponse,
	std::string& outError,
	int timeoutMs = ToolExecutionTimeouts::kDefaultResponseMs,
	int connectTimeoutMs = ToolExecutionTimeouts::kConnectMs);

// 验证本机 MCP 实例端口及实例标识
bool ProbeInstance(
	int port,
	const std::string& expectedInstanceId,
	std::string& outError,
	int timeoutMs = 500);

// 构建回环转发协议的内部自测报告
std::string BuildSelfTestReportJson();

} // namespace LocalMcpProxyTransport
