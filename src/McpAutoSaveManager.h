// McpAutoSaveManager.h
// 管理 MCP 工程写入成功后的可选自动保存。
#pragma once

#include <string>
#include <string_view>

class AIJsonConfig;

namespace McpAutoSaveManager {

inline constexpr std::string_view kConfigKey = "mcp_auto_save_after_write";

// 单次自动保存尝试的结构化结果。
struct Result {
	bool enabled = false;
	bool eligible = false;
	bool attempted = false;
	bool saved = false;
	std::string sourceFilePathLocal;
	std::string reason;
};

// 未配置时返回 false，确保升级后默认不开启。
bool IsEnabled(const AIJsonConfig& config);

// 仅为成功修改当前工程的 MCP 工具保存已存在于磁盘的 .e/.ec 源文件。
Result TrySaveAfterSuccessfulTool(
	const AIJsonConfig& config,
	std::string_view toolName,
	bool toolSucceeded,
	bool projectChanged);

} // namespace McpAutoSaveManager
