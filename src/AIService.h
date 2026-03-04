#pragma once

#include <functional>
#include <string>
#include <vector>

class ConfigManager;

enum class AITaskKind {
	OptimizeFunction,
	AddCommentsToFunction,
	TranslateFunctionAndVariables,
	TranslateText,
	AddByCurrentPageType
};

struct AISettings {
	std::string baseUrl;
	std::string apiKey;
	std::string model;
	std::string extraSystemPrompt;
	int timeoutMs = 120000;
	double temperature = 0.2;
};

struct AIResult {
	bool ok = false;
	std::string content;
	std::string error;
	int httpStatus = 0;
};

struct AIChatMessage {
	std::string role;   // "system" | "user" | "assistant"
	std::string content;
};

struct AIChatToolEvent {
	std::string name;
	std::string argumentsJson;
	std::string resultJson;
	bool ok = false;
};

struct AIChatResult {
	bool ok = false;
	std::string content;
	std::string error;
	int httpStatus = 0;
	std::vector<AIChatToolEvent> toolEvents;
};

class AIService {
public:
	static bool LoadSettings(ConfigManager& config, AISettings& outSettings);
	static void SaveSettings(ConfigManager& config, const AISettings& settings);
	static bool HasRequiredSettings(const AISettings& settings, std::string& outMissingField);
	static std::string BuildTaskDisplayName(AITaskKind kind);
	static AIResult ExecuteTask(AITaskKind kind, const std::string& inputText, const AISettings& settings);
	static AIChatResult ExecuteChatWithTools(
		const std::vector<AIChatMessage>& contextMessages,
		const AISettings& settings,
		const std::function<std::string(const std::string& toolName, const std::string& argumentsJson, bool& outOk)>& toolCallback);
	static std::string NormalizeModelOutputToCode(const std::string& modelText);
	static std::string Trim(const std::string& text);

private:
	static std::string BuildEndpoint(const std::string& baseUrl);
	static std::string BuildSystemPrompt(AITaskKind kind, const AISettings& settings);
};
