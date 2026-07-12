#pragma once

#include <filesystem>
#include <string>
#include <vector>

#include "..\\thirdparty\\json.hpp"

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
	// 来自 mcpServers 段（兼容其他客户端的配置），只读：不由本程序回写，
	// 配置界面中不可编辑。
	bool readOnly = false;
};

// MCP 工具自动允许授权。当前按 serverId + toolName + schemaHash 精确匹配；
// 旧版本的 "*"/"*" 整服授权不会再被承认。
struct AIChatMcpApprovalGrant {
	std::string serverId;
	std::string toolName;
	std::string schemaHash;
	long long createdAtUnixMs = 0;
	long long updatedAtUnixMs = 0;
};

// 当前 MCP 配置格式版本。
inline constexpr int kAIChatMcpConfigVersion = 2;

// AI 对话 MCP 客户端配置。
struct AIChatMcpConfig {
	int version = kAIChatMcpConfigVersion;
	std::vector<AIChatMcpServerConfig> servers;
	std::vector<AIChatMcpApprovalGrant> approvalGrants;
	// mcpServers 段原样保留（兼容其他客户端），序列化时原样写回，
	// 避免只读服务器在保存后丢失。空表示没有该段。
	nlohmann::json mcpServersRaw;
};

namespace AIChatMcpConfigStore {

// 获取 MCP 配置文件路径。
std::filesystem::path GetConfigPath();
// 构建不带服务器的默认配置 JSON。
std::string BuildDefaultConfigJson();
// 解析 MCP 配置 JSON。
bool ParseConfigJson(const std::string& jsonText, AIChatMcpConfig& outConfig, std::string& outError);
// 序列化 MCP 配置 JSON（磁盘格式：只读服务器写回 mcpServers 段）。
std::string SerializeConfigJson(const AIChatMcpConfig& config, bool pretty = true);
// 序列化供配置界面展示的 JSON：所有服务器（含 mcpServers 只读项）都放入
// servers 数组并带 read_only 标记，便于界面统一渲染，不写回磁盘。
std::string SerializeConfigForUi(const AIChatMcpConfig& config);
// 读取 MCP 配置文件。
bool Load(AIChatMcpConfig& outConfig, std::string* outError = nullptr);
// 保存 MCP 配置文件。
bool Save(const AIChatMcpConfig& config, std::string* outError = nullptr);
// 保存已经校验过的配置 JSON 文本。
// preserveMissingMcpServers=true 时（供只编辑 servers 数组的 WebView 使用），
// 若来文缺少 mcpServers 段则从磁盘保留旧值；false（默认，供原生 JSON 编辑器
// 使用）则原样写入来文，允许用户真正删除 mcpServers 段。
bool SaveJsonText(const std::string& jsonText, std::string* outError = nullptr, bool preserveMissingMcpServers = false);
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
