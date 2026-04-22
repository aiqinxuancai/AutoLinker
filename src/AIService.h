#pragma once

#include <functional>
#include <string>
#include <vector>

class AIJsonConfig;
class ConfigManager;
class HttpRequestCancellation;

// AI 任务类型。
enum class AITaskKind {
	OptimizeFunction,
	AddCommentsToFunction,
	TranslateFunctionAndVariables,
	TranslateText,
	AddByCurrentPageType
};

// AI 协议类型。
enum class AIProtocolType {
	OpenAI = 0,
	Gemini = 1,
	Claude = 2,
	OpenAIResponses = 3
};

// AI 设置。
struct AISettings {
	AIProtocolType protocolType = AIProtocolType::OpenAI;
	std::string baseUrl;
	std::string apiKey;
	std::string model;
	std::string extraSystemPrompt;
	std::string tavilyApiKey;
	int timeoutMs = 120000;
	// 工具调用最大轮数。
	int maxToolRounds = 48;
	double temperature = 0.2;
};

// AI 单次任务结果。
struct AIResult {
	bool ok = false;
	std::string content;
	std::string error;
	int httpStatus = 0;
};

// AI 对话消息。
struct AIChatMessage {
	std::string role;   // "system" | "user" | "assistant"
	std::string content;
};

// AI 工具调用事件。
struct AIChatToolEvent {
	std::string name;
	std::string argumentsJson;
	std::string resultJson;
	bool ok = false;
};

// AI 对话结果。
struct AIChatResult {
	bool ok = false;
	bool cancelled = false;
	std::string content;
	std::string error;
	int httpStatus = 0;
	std::vector<AIChatToolEvent> toolEvents;
};

class AIService {
public:
	// 从 JSON 配置加载 AI 设置，若 JSON 无数据且 iniConfig 非空则自动从 INI 迁移。
	static bool LoadSettings(AIJsonConfig& jsonConfig, ConfigManager* iniConfig, AISettings& outSettings);
	// 将 AI 设置保存到 JSON 配置。
	static void SaveSettings(AIJsonConfig& jsonConfig, const AISettings& settings);
	static bool HasRequiredSettings(const AISettings& settings, std::string& outMissingField);
	static AIProtocolType ParseProtocolType(const std::string& text);
	static std::string ProtocolTypeToString(AIProtocolType protocolType);
	static std::string ProtocolTypeDisplayName(AIProtocolType protocolType);
	static std::string BuildTaskDisplayName(AITaskKind kind);
	static AIResult ExecuteTask(AITaskKind kind, const std::string& inputText, const AISettings& settings);
	static AIChatResult ExecuteChatWithTools(
		const std::vector<AIChatMessage>& contextMessages,
		const AISettings& settings,
		const std::function<std::string(const std::string& toolName, const std::string& argumentsJson, bool& outOk)>& toolCallback,
		const std::function<void(const std::string& deltaText)>& streamCallback = {},
		const std::function<bool()>& cancelCallback = {},
		HttpRequestCancellation* cancelContext = nullptr);
	static std::string BuildPublicToolCatalogJson();
	static std::string NormalizeModelOutputToCode(const std::string& modelText);
	static std::string Trim(const std::string& text);

private:
	static std::string BuildEndpoint(const std::string& baseUrl);
	static std::string BuildSystemPrompt(AITaskKind kind, const AISettings& settings);
};
