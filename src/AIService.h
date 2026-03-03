#pragma once

#include <string>

class ConfigManager;

enum class AITaskKind {
	OptimizeFunction,
	AddCommentsToFunction,
	TranslateFunctionAndVariables,
	TranslateText,
	CompleteApiDeclarations,
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

class AIService {
public:
	static bool LoadSettings(ConfigManager& config, AISettings& outSettings);
	static void SaveSettings(ConfigManager& config, const AISettings& settings);
	static bool HasRequiredSettings(const AISettings& settings, std::string& outMissingField);
	static std::string BuildTaskDisplayName(AITaskKind kind);
	static AIResult ExecuteTask(AITaskKind kind, const std::string& inputText, const AISettings& settings);
	static std::string NormalizeModelOutputToCode(const std::string& modelText);
	static std::string Trim(const std::string& text);

private:
	static std::string BuildEndpoint(const std::string& baseUrl);
	static std::string BuildSystemPrompt(AITaskKind kind, const AISettings& settings);
};
