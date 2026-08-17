#pragma once

#include <functional>
#include <string>
#include <vector>

#include "AIImageAttachment.h"

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

// AI 思考等级。
enum class AIThinkingLevel {
	Off = 0,
	Low = 1,
	Medium = 2,
	High = 3,
	XHigh = 4,
	Max = 5,
	Ultra = 6
};

// AI 源码编辑基准模式。
enum class AISourceEditMode {
	RealPageFirst = 0,
	MirrorSourceBase = 1
};

// AI 图片输入能力覆盖模式。
enum class AIImageInputMode {
	Auto = 0,
	Enabled = 1,
	Disabled = 2
};

// 单个 AI API 端点设置。
struct AIEndpointSettings {
	std::string endpointId;
	std::string endpointName;
	AIProtocolType protocolType = AIProtocolType::OpenAI;
	AIThinkingLevel thinkingLevel = AIThinkingLevel::Off;
	std::string baseUrl;
	std::string apiKey;
	std::string model;
	std::string extraSystemPrompt;
	std::string customHeadersText;
	int timeoutMs = 120000;
	double temperature = 0.2;
	int contextWindowTokens = 0; // 0 = 未设置，回落到模型表/默认
	AIImageInputMode imageInputMode = AIImageInputMode::Auto;
	int retryCount = 5; // 首次请求之外的额外重试次数
};

// AI 设置，包含全局选项和当前活动目标解析出的有序端点。
struct AISettings : AIEndpointSettings {
	AISourceEditMode sourceEditMode = AISourceEditMode::RealPageFirst;
	std::string tavilyApiKey;
	std::vector<AIEndpointSettings> endpointCandidates;
	std::string activeTargetType = "endpoint";
	std::string activeTargetId;
};

// AI 单次任务结果。
struct AIResult {
	bool ok = false;
	bool endpointEstablished = false;
	std::string content;
	std::string error;
	std::string endpointId;
	std::string endpointName;
	int httpStatus = 0;
};

// AI 对话消息。
struct AIChatMessage {
	std::string role;   // "system" | "user" | "assistant"
	std::string content;
	std::string reasoningContent;
	std::string rawMessageJsonUtf8;
	std::vector<AIImageAttachment> attachments;
};

// AI 工具调用事件。
struct AIChatToolEvent {
	std::string name;
	std::string argumentsJson;
	std::string resultJson;
	bool ok = false;
};

// AI 长任务终止原因。
enum class AIChatRunTerminationReason {
	None = 0,
	Completed,
	Cancelled,
	ProviderError,
	Stalled,
	CompactionFailed
};

// AI 长任务中尚未完成或刚完成的工具调用。
struct AIChatCheckpointToolCall {
	std::string callId;
	std::string name;
	std::string argumentsJson;
	std::string resultJson;
	bool completed = false;
	bool ok = false;
};

// AI 长任务的可恢复检查点。
struct AIChatRunCheckpoint {
	int schemaVersion = 1;
	AIProtocolType protocolType = AIProtocolType::OpenAI;
	std::string model;
	std::string state = "running"; // running | paused | completed
	std::string summary;
	int samplingRounds = 0;
	int compactionCount = 0;
	int promptTokens = 0;
	int totalTokens = 0;
	bool hasUsage = false;
	std::vector<AIChatMessage> contextMessages;
	std::vector<AIChatCheckpointToolCall> toolCalls;
};

// AI 长任务运行选项。
struct AIChatRunOptions {
	const AIChatRunCheckpoint* resumeCheckpoint = nullptr;
	std::function<void(const AIChatRunCheckpoint& checkpoint)> checkpointCallback;
	// 仅在计划阶段向内置聊天公开结构化询问工具。
	bool enablePlanUserInput = false;
	// 仅在活动 Goal 中公开查询和结束 Goal 的工具。
	bool enableGoalTools = false;
	// 在模型请求之间的安全边界取出新用户输入；参数为刚完成的助手回复。
	std::function<std::vector<AIChatMessage>(const std::string& completedAssistantContent)> takePendingUserInputsCallback;
	// 流式协议重连前清理当前未完成的界面预览，避免重放增量造成重复显示。
	std::function<void()> streamRetryCallback;
};

// AI 对话结果。
struct AIChatResult {
	bool ok = false;
	bool cancelled = false;
	bool endpointEstablished = false;
	// 兼容旧调试接口；长期任务不再按固定工具轮数终止。
	bool toolRoundsExceeded = false;
	bool paused = false;
	AIChatRunTerminationReason terminationReason = AIChatRunTerminationReason::None;
	std::string content;
	std::string reasoningContent;
	std::string error;
	std::string endpointId;
	std::string endpointName;
	int httpStatus = 0;
	std::vector<AIChatToolEvent> toolEvents;
	std::vector<std::string> contextPrefixRawMessagesUtf8;
	// 本轮真实 token 用量（用于上下文压缩触发判定）。
	bool hasUsage = false;
	int promptTokens = 0; // 输入 token —— 衡量「上下文有多满」的关键数
	int totalTokens = 0;  // prompt+completion，仅日志诊断用
	int samplingRounds = 0;
	int compactionCount = 0;
	bool hasCheckpoint = false;
	AIChatRunCheckpoint checkpoint;
	std::vector<AIChatMessage> continuationMessages;
};

class AIService {
public:
	// 从 JSON 配置加载 AI 设置，若 JSON 无数据且 iniConfig 非空则自动从 INI 迁移。
	static bool LoadSettings(AIJsonConfig& jsonConfig, ConfigManager* iniConfig, AISettings& outSettings);
	// 将 AI 设置保存到 JSON 配置。
	static void SaveSettings(AIJsonConfig& jsonConfig, const AISettings& settings);
	static bool HasRequiredSettings(const AISettings& settings, std::string& outMissingField);
	// 解析当前模型的有效上下文窗口（token）。优先级：用户配置 > 内置模型表 > 默认 200000。
	static int ResolveContextWindowTokens(const AISettings& settings);
	static AIProtocolType ParseProtocolType(const std::string& text);
	static std::string ProtocolTypeToString(AIProtocolType protocolType);
	static std::string ProtocolTypeDisplayName(AIProtocolType protocolType);
	static AIThinkingLevel ParseThinkingLevel(const std::string& text);
	static std::string ThinkingLevelToString(AIThinkingLevel thinkingLevel);
	static std::string ThinkingLevelDisplayName(AIThinkingLevel thinkingLevel);
	static AISourceEditMode ParseSourceEditMode(const std::string& text);
	static std::string SourceEditModeToString(AISourceEditMode mode);
	static std::string SourceEditModeDisplayName(AISourceEditMode mode);
	static AIImageInputMode ParseImageInputMode(const std::string& text);
	static std::string ImageInputModeToString(AIImageInputMode mode);
	static std::string ImageInputModeDisplayName(AIImageInputMode mode);
	// 结合协议、内置模型能力表和用户覆盖，判断当前配置是否允许图片输入。
	static bool SupportsImageInput(const AISettings& settings);
	// 校验自定义请求头多行文本格式。
	static bool ValidateCustomHeadersText(const std::string& headerText, std::string& outError);
	// 测试当前 AI 配置的接口连通性。
	static AIResult TestConnection(const AISettings& settings);
	static std::string BuildTaskDisplayName(AITaskKind kind);
	static AIResult ExecuteTask(AITaskKind kind, const std::string& inputText, const AISettings& settings);
	static AIChatResult ExecuteChatWithTools(
		const std::vector<AIChatMessage>& contextMessages,
		const AISettings& settings,
		const std::function<std::string(const std::string& toolName, const std::string& argumentsJson, bool& outOk)>& toolCallback,
		const std::function<void(const std::string& deltaText)>& streamCallback = {},
		const std::function<bool()>& cancelCallback = {},
		HttpRequestCancellation* cancelContext = nullptr,
		const AIChatRunOptions& runOptions = {});
	static std::string BuildPublicToolCatalogJson();
	// 构建对外 MCP（本地端口）initialize 返回的 instructions：告知外部客户端
	// 必须经本 MCP 工具读写易语言工程，并附带最易出错的易语言语法约定浓缩版。
	static std::string BuildExternalMcpInstructions();
	// 构建 Agent 工具优化与 Responses 流式解析的内部自测报告。
	static std::string BuildAgentOptimizationSelfTestJson();
	// 构建四种协议图片消息序列化与能力判断的内部自测报告。
	static std::string BuildImageInputSelfTestJson();
	// 构建 API 端点配置迁移、分组顺序与重试范围的内部自测报告。
	static std::string BuildEndpointConfigSelfTestJson();
	static std::string NormalizeModelOutputToCode(const std::string& modelText);
	static std::string Trim(const std::string& text);

private:
	static AIResult TestConnectionSingle(const AISettings& settings);
	static AIResult ExecuteTaskSingle(AITaskKind kind, const std::string& inputText, const AISettings& settings);
	static AIChatResult ExecuteChatWithToolsSingle(
		const std::vector<AIChatMessage>& contextMessages,
		const AISettings& settings,
		const std::function<std::string(const std::string& toolName, const std::string& argumentsJson, bool& outOk)>& toolCallback,
		const std::function<void(const std::string& deltaText)>& streamCallback,
		const std::function<bool()>& cancelCallback,
		HttpRequestCancellation* cancelContext,
		const AIChatRunOptions& runOptions);
	static std::string BuildEndpoint(const std::string& baseUrl);
	static std::string BuildSystemPrompt(AITaskKind kind, const AISettings& settings);
};
