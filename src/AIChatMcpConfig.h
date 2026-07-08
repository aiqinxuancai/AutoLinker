#pragma once

#include <filesystem>
#include <string>
#include <vector>

// MCP 请求头配置。
struct AIChatMcpHeaderConfig {
	std::string name;
	std::string value;
};

// MCP stdio 环境变量配置。
struct AIChatMcpEnvConfig {
	std::string name;
	std::string value;
};

// 外部 MCP 服务器配置。
struct AIChatMcpServerConfig {
	std::string id;
	std::string name;
	std::string transport = "streamable_http";
	std::string url;
	std::string command;
	std::vector<std::string> arguments;
	std::string workingDirectory;
	bool enabled = false;
	int timeoutMs = 120000;
	std::vector<AIChatMcpHeaderConfig> headers;
	std::vector<AIChatMcpEnvConfig> env;
};

// MCP 工具自动允许授权。
struct AIChatMcpApprovalGrant {
	std::string serverId;
	std::string toolName;
	std::string schemaHash;
	long long createdAtUnixMs = 0;
	long long updatedAtUnixMs = 0;
};

// AI 对话 MCP 客户端配置。
struct AIChatMcpConfig {
	int version = 1;
	std::vector<AIChatMcpServerConfig> servers;
	std::vector<AIChatMcpApprovalGrant> approvalGrants;
};

namespace AIChatMcpConfigStore {

// 获取 MCP 配置文件路径。
std::filesystem::path GetConfigPath();
// 构建带示例服务器的默认配置 JSON。
std::string BuildDefaultConfigJson();
// 解析 MCP 配置 JSON。
bool ParseConfigJson(const std::string& jsonText, AIChatMcpConfig& outConfig, std::string& outError);
// 序列化 MCP 配置 JSON。
std::string SerializeConfigJson(const AIChatMcpConfig& config, bool pretty = true);
// 读取 MCP 配置文件。
bool Load(AIChatMcpConfig& outConfig, std::string* outError = nullptr);
// 保存 MCP 配置文件。
bool Save(const AIChatMcpConfig& config, std::string* outError = nullptr);
// 保存已经校验过的配置 JSON 文本。
bool SaveJsonText(const std::string& jsonText, std::string* outError = nullptr);
// 判断指定 MCP 工具是否已有自动允许授权。
bool HasApprovalGrant(
	const AIChatMcpConfig& config,
	const std::string& serverId,
	const std::string& toolName,
	const std::string& schemaHash);
// 写入或刷新自动允许授权。
void UpsertApprovalGrant(
	AIChatMcpConfig& config,
	const std::string& serverId,
	const std::string& toolName,
	const std::string& schemaHash);

} // namespace AIChatMcpConfigStore
