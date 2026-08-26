#include "McpAutoSaveManager.h"

#include <algorithm>
#include <cctype>
#include <filesystem>
#include <format>

#include "AIChatToolRegistry.h"
#include "AIJsonConfig.h"
#include "Global.h"
#include "IDEFacade.h"
#include "Logger.h"

namespace McpAutoSaveManager {
namespace {

std::string TrimAsciiCopy(std::string text)
{
	const auto isSpace = [](unsigned char ch) { return std::isspace(ch) != 0; };
	text.erase(text.begin(), std::find_if_not(text.begin(), text.end(), isSpace));
	text.erase(std::find_if_not(text.rbegin(), text.rend(), isSpace).base(), text.end());
	return text;
}

bool HasSavedSourcePath(const std::string& sourcePathLocal)
{
	if (sourcePathLocal.empty()) {
		return false;
	}
	std::error_code error;
	const std::filesystem::path path(sourcePathLocal);
	if (!std::filesystem::is_regular_file(path, error) || error) {
		return false;
	}
	std::string extension = path.extension().string();
	std::transform(extension.begin(), extension.end(), extension.begin(), [](unsigned char ch) {
		return static_cast<char>(std::tolower(ch));
	});
	return extension == ".e" || extension == ".ec";
}

} // namespace

bool IsEnabled(const AIJsonConfig& config)
{
	const std::string value = config.getGlobalValue(std::string(kConfigKey));
	return value == "true" || value == "1";
}

Result TrySaveAfterSuccessfulTool(
	const AIJsonConfig& config,
	std::string_view toolName,
	bool toolSucceeded,
	bool projectChanged)
{
	Result result;
	result.enabled = IsEnabled(config);
	result.eligible = toolSucceeded && projectChanged && AIChatToolRegistry::ModifiesProject(toolName);
	if (!result.enabled || !result.eligible) {
		return result;
	}

	UpdateCurrentOpenSourceFile();
	result.sourceFilePathLocal = TrimAsciiCopy(g_nowOpenSourceFilePath);
	if (!HasSavedSourcePath(result.sourceFilePathLocal)) {
		result.reason = "source_file_not_saved_to_disk";
		Logger::Instance().WriteAndIde(
			"MCP",
			std::format("auto_save_skipped tool={} reason={}", toolName, result.reason));
		return result;
	}

	result.attempted = true;
	result.saved = IDEFacade::Instance().SaveFile();
	if (!result.saved) {
		result.reason = "FN_SAVE_FILE_failed";
	}
	Logger::Instance().WriteAndIde(
		"MCP",
		std::format(
			"auto_save_finished tool={} ok={} source={}",
			toolName,
			result.saved ? 1 : 0,
			result.sourceFilePathLocal));
	return result;
}

} // namespace McpAutoSaveManager
