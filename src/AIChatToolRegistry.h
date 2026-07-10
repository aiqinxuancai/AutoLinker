#pragma once

#include <string>
#include <string_view>

#include "..\\thirdparty\\json.hpp"

// AI 工具注册元数据，用于统一控制工具可见性、刷新门禁和风险属性。
namespace AIChatToolRegistry {

struct ToolMetadata {
	std::string_view name;
	bool externalPublic = false;
	bool dependencyManagement = false;
	bool requiresWorkspaceRefresh = false;
	bool destructive = false;
	bool interactive = false;
};

// 查找一个 AutoLinker 原生工具；外部 MCP 动态工具不属于此注册表。
const ToolMetadata* Find(std::string_view toolName);

// 判断工具是否允许由本地 MCP 对外公开调用。
bool IsExternalPublic(std::string_view toolName);

// 判断工具是否为按用户明确意图开放的依赖管理工具。
bool IsDependencyManagement(std::string_view toolName);

// 判断外部 MCP 会话在调用该工具前是否必须刷新工作区镜像。
bool RequiresWorkspaceRefresh(std::string_view toolName);

// 按注册表过滤工具目录，仅保留真正允许对外公开的原生工具。
nlohmann::json FilterExternalPublicCatalog(const nlohmann::json& catalog);

// 按工具 inputSchema 校验调用参数，失败时返回稳定、可读的错误说明。
bool ValidateArguments(
	const nlohmann::json& arguments,
	const nlohmann::json& inputSchema,
	std::string& outError);

} // namespace AIChatToolRegistry
