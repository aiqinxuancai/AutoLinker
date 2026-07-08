#pragma once

#include <functional>
#include <string>
#include <vector>

#include "..\\thirdparty\\json.hpp"

#include "AIChatMcpConfig.h"

class HttpRequestCancellation;

// MCP 工具元数据。
struct AIChatMcpToolInfo {
	std::string serverId;
	std::string serverName;
	std::string originalName;
	std::string modelName;
	std::string description;
	std::string schemaHash;
	nlohmann::json inputSchema = nlohmann::json::object();
};

// MCP 工具执行确认上下文。
struct AIChatMcpApprovalContext {
	std::string serverId;
	std::string serverName;
	std::string toolName;
	std::string modelToolName;
	std::string schemaHash;
	std::string argumentsJsonUtf8;
};

// MCP 工具执行结果。
struct AIChatMcpExecutionResult {
	bool ok = false;
	bool denied = false;
	bool cancelled = false;
	int httpStatus = 0;
	std::string resultJsonLocal;
	std::string errorUtf8;
};

namespace AIChatMcpClient {

// 判断工具名是否为 AutoLinker 生成的 MCP 工具名。
bool IsMcpModelToolName(const std::string& toolName);
// 获取当前可用外部 MCP 工具目录。
std::vector<AIChatMcpToolInfo> LoadEnabledTools();
// 将外部 MCP 工具追加到 catalog。
nlohmann::json AppendMcpToolsToCatalog(const nlohmann::json& baseCatalog);
// 执行一个外部 MCP 工具。
AIChatMcpExecutionResult ExecuteTool(
	const std::string& modelToolName,
	const std::string& argumentsJsonUtf8,
	const std::function<bool(const AIChatMcpApprovalContext& context, bool& outAutoAllow)>& approvalCallback,
	const std::function<bool()>& cancelCallback = {},
	HttpRequestCancellation* cancellation = nullptr);
// 构造无 IDE 依赖的自检报告。
std::string BuildSelfTestReportJson();

} // namespace AIChatMcpClient
