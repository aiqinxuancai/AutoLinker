#include "AutoLinkerTestApi.h"

#include <algorithm>
#include <cstring>
#include <string>
#include <vector>

#include "..\\thirdparty\\json.hpp"

#include "AIChatTooling.h"
#include "AIService.h"
#include "AutoLinkerVersion.h"
#include "GameAnalyticsClient.h"
#include "PathHelper.h"
#include "Version.h"

namespace {

int CopyStringToBuffer(const std::string& value, char* buffer, int bufferSize)
{
	if (buffer == nullptr || bufferSize <= 0) {
		return AUTOLINKER_TEST_STRING_INVALID_ARGUMENT;
	}

	const size_t requiredSize = value.size() + 1;
	if (requiredSize > static_cast<size_t>(bufferSize)) {
		return AUTOLINKER_TEST_STRING_BUFFER_TOO_SMALL;
	}

	std::memcpy(buffer, value.c_str(), requiredSize);
	return static_cast<int>(value.size());
}

std::string BuildDeepSeekToolArgumentsJson(const std::string& toolName)
{
	if (toolName == "fetch_url") {
		return R"({"url":"https://api-docs.deepseek.com/quick_start/rate_limit","timeout_seconds":30,"max_bytes":262144})";
	}
	if (toolName == "extract_web_document") {
		return R"({"url":"https://api-docs.deepseek.com/guides/thinking_mode","timeout_seconds":30,"max_bytes":262144})";
	}
	return "{}";
}

nlohmann::json BuildDeepSeekIntegrationResultJson(const AISettings& settings)
{
	return {
		{"provider", "deepseek"},
		{"model", settings.model},
		{"base_url", settings.baseUrl},
		{"thinking_level", AIService::ThinkingLevelToString(settings.thinkingLevel)},
		{"protocol", AIService::ProtocolTypeToString(settings.protocolType)}
	};
}

std::string BuildOpenAIToolArgumentsJson(const std::string& toolName)
{
	if (toolName == "fetch_url") {
		return R"({"url":"https://developers.openai.com/api/docs/api-reference/chat/create-chat-completion","timeout_seconds":30,"max_bytes":262144})";
	}
	if (toolName == "extract_web_document") {
		return R"({"url":"https://developers.openai.com/api/docs/api-reference/responses/create","timeout_seconds":30,"max_bytes":262144})";
	}
	return "{}";
}

nlohmann::json BuildOpenAIIntegrationResultJson(const AISettings& settings)
{
	return {
		{"provider", "openai"},
		{"model", settings.model},
		{"base_url", settings.baseUrl},
		{"thinking_level", AIService::ThinkingLevelToString(settings.thinkingLevel)},
		{"protocol", AIService::ProtocolTypeToString(settings.protocolType)}
	};
}

std::string BuildGeminiToolArgumentsJson(const std::string& toolName)
{
	if (toolName == "fetch_url") {
		return R"({"url":"https://ai.google.dev/gemini-api/docs/models/gemini","timeout_seconds":30,"max_bytes":4096})";
	}
	if (toolName == "extract_web_document") {
		return R"({"url":"https://ai.google.dev/gemini-api/docs/function-calling","timeout_seconds":30,"max_bytes":4096})";
	}
	return "{}";
}

nlohmann::json BuildGeminiIntegrationResultJson(const AISettings& settings)
{
	return {
		{"provider", "gemini"},
		{"model", settings.model},
		{"base_url", settings.baseUrl},
		{"thinking_level", AIService::ThinkingLevelToString(settings.thinkingLevel)},
		{"protocol", AIService::ProtocolTypeToString(settings.protocolType)}
	};
}

std::string BuildClaudeToolArgumentsJson(const std::string& toolName)
{
	if (toolName == "fetch_url") {
		return R"({"url":"https://docs.anthropic.com/en/docs/about-claude/models/overview","timeout_seconds":30,"max_bytes":4096})";
	}
	if (toolName == "extract_web_document") {
		return R"({"url":"https://docs.anthropic.com/en/docs/agents-and-tools/tool-use/overview","timeout_seconds":30,"max_bytes":4096})";
	}
	return "{}";
}

nlohmann::json BuildClaudeIntegrationResultJson(const AISettings& settings)
{
	return {
		{"provider", "claude"},
		{"model", settings.model},
		{"base_url", settings.baseUrl},
		{"thinking_level", AIService::ThinkingLevelToString(settings.thinkingLevel)},
		{"protocol", AIService::ProtocolTypeToString(settings.protocolType)}
	};
}

std::string ExtractMessageContentText(const nlohmann::json& parsed)
{
	if (!parsed.is_object() || !parsed.contains("content")) {
		return std::string();
	}
	if (parsed["content"].is_string()) {
		return parsed["content"].get<std::string>();
	}
	if (!parsed["content"].is_array()) {
		return std::string();
	}

	std::string content;
	for (const auto& item : parsed["content"]) {
		if (!item.is_object()) {
			continue;
		}
		const std::string contentType = item.value("type", std::string());
		if ((contentType == "output_text" || contentType == "text") &&
			item.contains("text") &&
			item["text"].is_string()) {
			content += item["text"].get<std::string>();
		}
	}
	return content;
}

std::vector<AIChatMessage> BuildFollowupMessagesFromChatResult(
	const std::vector<AIChatMessage>& prefixMessages,
	const AIChatResult& toolChatResult)
{
	std::vector<AIChatMessage> followupMessages = prefixMessages;
	for (const auto& rawMessageJsonUtf8 : toolChatResult.contextPrefixRawMessagesUtf8) {
		nlohmann::json parsed;
		try {
			parsed = nlohmann::json::parse(rawMessageJsonUtf8);
		}
		catch (...) {
			continue;
		}
		if (!parsed.is_object()) {
			continue;
		}

		const std::string role = parsed.value("role", std::string());
		const std::string type = parsed.value("type", std::string());
		if (role == "assistant") {
			followupMessages.push_back({
				"assistant",
				ExtractMessageContentText(parsed),
				parsed.value("reasoning_content", std::string()),
				rawMessageJsonUtf8
			});
		}
		else if (role == "tool") {
			followupMessages.push_back({
				"tool",
				ExtractMessageContentText(parsed),
				"",
				rawMessageJsonUtf8
			});
		}
		else if (!type.empty()) {
			followupMessages.push_back({
				"tool",
				ExtractMessageContentText(parsed),
				"",
				rawMessageJsonUtf8
			});
		}
	}
	return followupMessages;
}

std::string DumpJsonPrettySafe(const nlohmann::json& value)
{
	return value.dump(2, ' ', false, nlohmann::json::error_handler_t::replace);
}

int RunOpenAIIntegrationTestInternal(
	AIProtocolType protocolType,
	const char* apiKey,
	const char* model,
	const char* baseUrl,
	char* buffer,
	int bufferSize)
{
	if (apiKey == nullptr || model == nullptr) {
		return AUTOLINKER_TEST_STRING_INVALID_ARGUMENT;
	}

	const char* defaultBaseUrl = "https://api.openai.com/v1";
	std::string step = "init";
	try {
		AISettings settings = {};
		settings.protocolType = protocolType;
		settings.thinkingLevel = AIThinkingLevel::High;
		settings.baseUrl = (baseUrl != nullptr && baseUrl[0] != '\0') ? baseUrl : defaultBaseUrl;
		settings.apiKey = apiKey;
		settings.model = model;
		settings.timeoutMs = 180000;
		settings.temperature = 0;

		nlohmann::json report = BuildOpenAIIntegrationResultJson(settings);
		report["step"] = step;

		step = "test_connection";
		report["step"] = step;
		const AIResult connectionResult = AIService::TestConnection(settings);
		report["test_connection"] = {
			{"ok", connectionResult.ok},
			{"http_status", connectionResult.httpStatus},
			{"content", connectionResult.content},
			{"error", connectionResult.error}
		};
		if (!connectionResult.ok) {
			report["ok"] = false;
			return CopyStringToBuffer(DumpJsonPrettySafe(report), buffer, bufferSize);
		}

		step = "simple_task";
		report["step"] = step;
		const AIResult simpleTaskResult = AIService::ExecuteTask(
			AITaskKind::TranslateText,
			"只返回这四个字符：测试通过",
			settings);
		report["simple_task"] = {
			{"ok", simpleTaskResult.ok},
			{"http_status", simpleTaskResult.httpStatus},
			{"content", simpleTaskResult.content},
			{"error", simpleTaskResult.error}
		};
		if (!simpleTaskResult.ok) {
			report["ok"] = false;
			return CopyStringToBuffer(DumpJsonPrettySafe(report), buffer, bufferSize);
		}

		step = "tool_chat";
		report["step"] = step;
		std::vector<AIChatMessage> contextMessages;
		contextMessages.push_back({
			"user",
			"你必须先后调用两个工具：先 fetch_url 读取 https://developers.openai.com/api/docs/api-reference/chat/create-chat-completion ，再 extract_web_document 读取 https://developers.openai.com/api/docs/api-reference/responses/create 。完成后仅用一行中文回答，格式必须是：Chat页已读；Responses页已读；工具数=N。",
			"",
			""
		});

		std::vector<std::string> streamedDeltas;
		const AIChatResult toolChatResult = AIService::ExecuteChatWithTools(
			contextMessages,
			settings,
			[](const std::string& toolName, const std::string&, bool& outOk) -> std::string {
				const std::string actualArgs = BuildOpenAIToolArgumentsJson(toolName);
				return ExecuteToolCall(toolName, actualArgs, outOk, false);
			},
			[&streamedDeltas](const std::string& deltaText) {
				if (!deltaText.empty()) {
					streamedDeltas.push_back(deltaText);
				}
			});

		nlohmann::json toolEvents = nlohmann::json::array();
		for (const auto& evt : toolChatResult.toolEvents) {
			toolEvents.push_back({
				{"name", evt.name},
				{"arguments_json", evt.argumentsJson},
				{"result_json", evt.resultJson},
				{"ok", evt.ok}
			});
		}
		report["tool_chat"] = {
			{"ok", toolChatResult.ok},
			{"cancelled", toolChatResult.cancelled},
			{"http_status", toolChatResult.httpStatus},
			{"content", toolChatResult.content},
			{"reasoning_content_present", !toolChatResult.reasoningContent.empty()},
			{"reasoning_content_size", toolChatResult.reasoningContent.size()},
			{"error", toolChatResult.error},
			{"tool_events", std::move(toolEvents)},
			{"stream_chunk_count", streamedDeltas.size()},
			{"hidden_context_message_count", toolChatResult.contextPrefixRawMessagesUtf8.size()}
		};
		if (!toolChatResult.ok) {
			report["ok"] = false;
			return CopyStringToBuffer(DumpJsonPrettySafe(report), buffer, bufferSize);
		}

		step = "followup_chat";
		report["step"] = step;
		std::vector<AIChatMessage> followupMessages = BuildFollowupMessagesFromChatResult(contextMessages, toolChatResult);
		followupMessages.push_back({
			"assistant",
			toolChatResult.content,
			toolChatResult.reasoningContent,
			""
		});
		followupMessages.push_back({
			"user",
			"只回答：上一轮你实际调用了几个工具？输出阿拉伯数字。",
			"",
			""
		});

		const AIChatResult followupResult = AIService::ExecuteChatWithTools(
			followupMessages,
			settings,
			[](const std::string& toolName, const std::string&, bool& outOk) -> std::string {
				const std::string actualArgs = BuildOpenAIToolArgumentsJson(toolName);
				return ExecuteToolCall(toolName, actualArgs, outOk, false);
			});
		report["followup_chat"] = {
			{"ok", followupResult.ok},
			{"cancelled", followupResult.cancelled},
			{"http_status", followupResult.httpStatus},
			{"content", followupResult.content},
			{"reasoning_content_present", !followupResult.reasoningContent.empty()},
			{"reasoning_content_size", followupResult.reasoningContent.size()},
			{"error", followupResult.error},
			{"tool_event_count", followupResult.toolEvents.size()}
		};

		report["ok"] =
			connectionResult.ok &&
			simpleTaskResult.ok &&
			toolChatResult.ok &&
			followupResult.ok &&
			toolChatResult.toolEvents.size() >= 2;
		return CopyStringToBuffer(DumpJsonPrettySafe(report), buffer, bufferSize);
	}
	catch (const std::exception& ex) {
		nlohmann::json report = {
			{"ok", false},
			{"provider", "openai"},
			{"model", model},
			{"base_url", (baseUrl != nullptr && baseUrl[0] != '\0') ? baseUrl : defaultBaseUrl},
			{"protocol", protocolType == AIProtocolType::OpenAIResponses ? "openai_responses" : "openai"},
			{"step", step},
			{"error", std::string("exception: ") + ex.what()}
		};
		return CopyStringToBuffer(DumpJsonPrettySafe(report), buffer, bufferSize);
	}
	catch (...) {
		nlohmann::json report = {
			{"ok", false},
			{"provider", "openai"},
			{"model", model},
			{"base_url", (baseUrl != nullptr && baseUrl[0] != '\0') ? baseUrl : defaultBaseUrl},
			{"protocol", protocolType == AIProtocolType::OpenAIResponses ? "openai_responses" : "openai"},
			{"step", step},
			{"error", "unknown exception"}
		};
		return CopyStringToBuffer(DumpJsonPrettySafe(report), buffer, bufferSize);
	}
}

int RunGeminiIntegrationTestInternal(
	const char* apiKey,
	const char* model,
	const char* baseUrl,
	char* buffer,
	int bufferSize)
{
	if (apiKey == nullptr || model == nullptr) {
		return AUTOLINKER_TEST_STRING_INVALID_ARGUMENT;
	}

	const char* defaultBaseUrl = "https://generativelanguage.googleapis.com";
	std::string step = "init";
	try {
		AISettings settings = {};
		settings.protocolType = AIProtocolType::Gemini;
		settings.thinkingLevel = AIThinkingLevel::Off;
		settings.baseUrl = (baseUrl != nullptr && baseUrl[0] != '\0') ? baseUrl : defaultBaseUrl;
		settings.apiKey = apiKey;
		settings.model = model;
		settings.timeoutMs = 180000;
		settings.temperature = 0;

		nlohmann::json report = BuildGeminiIntegrationResultJson(settings);
		report["step"] = step;

		step = "test_connection";
		report["step"] = step;
		const AIResult connectionResult = AIService::TestConnection(settings);
		report["test_connection"] = {
			{"ok", connectionResult.ok},
			{"http_status", connectionResult.httpStatus},
			{"content", connectionResult.content},
			{"error", connectionResult.error}
		};
		if (!connectionResult.ok) {
			report["ok"] = false;
			return CopyStringToBuffer(DumpJsonPrettySafe(report), buffer, bufferSize);
		}

		step = "simple_task";
		report["step"] = step;
		const AIResult simpleTaskResult = AIService::ExecuteTask(
			AITaskKind::TranslateText,
			"只返回这四个字符：测试通过",
			settings);
		report["simple_task"] = {
			{"ok", simpleTaskResult.ok},
			{"http_status", simpleTaskResult.httpStatus},
			{"content", simpleTaskResult.content},
			{"error", simpleTaskResult.error}
		};
		if (!simpleTaskResult.ok) {
			report["ok"] = false;
			return CopyStringToBuffer(DumpJsonPrettySafe(report), buffer, bufferSize);
		}

		step = "tool_chat";
		report["step"] = step;
		std::vector<AIChatMessage> contextMessages;
		contextMessages.push_back({
			"user",
			"你必须先后调用两个工具：先 fetch_url 读取 https://ai.google.dev/gemini-api/docs/models/gemini ，再 extract_web_document 读取 https://ai.google.dev/gemini-api/docs/function-calling 。完成后仅用一行中文回答，格式必须是：模型页已读；函数调用页已读；工具数=N。",
			"",
			""
		});

		std::vector<std::string> streamedDeltas;
		const AIChatResult toolChatResult = AIService::ExecuteChatWithTools(
			contextMessages,
			settings,
			[](const std::string& toolName, const std::string&, bool& outOk) -> std::string {
				const std::string actualArgs = BuildGeminiToolArgumentsJson(toolName);
				return ExecuteToolCall(toolName, actualArgs, outOk, false);
			},
			[&streamedDeltas](const std::string& deltaText) {
				if (!deltaText.empty()) {
					streamedDeltas.push_back(deltaText);
				}
			});

		nlohmann::json toolEvents = nlohmann::json::array();
		bool allToolEventsOk = true;
		for (const auto& evt : toolChatResult.toolEvents) {
			allToolEventsOk = allToolEventsOk && evt.ok;
			toolEvents.push_back({
				{"name", evt.name},
				{"arguments_json", evt.argumentsJson},
				{"result_json", evt.resultJson},
				{"ok", evt.ok}
			});
		}
		report["tool_chat"] = {
			{"ok", toolChatResult.ok},
			{"cancelled", toolChatResult.cancelled},
			{"http_status", toolChatResult.httpStatus},
			{"content", toolChatResult.content},
			{"reasoning_content_present", !toolChatResult.reasoningContent.empty()},
			{"reasoning_content_size", toolChatResult.reasoningContent.size()},
			{"error", toolChatResult.error},
			{"tool_events", std::move(toolEvents)},
			{"all_tool_events_ok", allToolEventsOk},
			{"stream_chunk_count", streamedDeltas.size()},
			{"hidden_context_message_count", toolChatResult.contextPrefixRawMessagesUtf8.size()}
		};
		if (!toolChatResult.ok) {
			report["ok"] = false;
			return CopyStringToBuffer(DumpJsonPrettySafe(report), buffer, bufferSize);
		}

		step = "followup_chat";
		report["step"] = step;
		std::vector<AIChatMessage> followupMessages = BuildFollowupMessagesFromChatResult(contextMessages, toolChatResult);
		followupMessages.push_back({
			"assistant",
			toolChatResult.content,
			toolChatResult.reasoningContent,
			""
		});
		followupMessages.push_back({
			"user",
			"只回答：上一轮你实际调用了几个工具？输出阿拉伯数字。",
			"",
			""
		});

		const AIChatResult followupResult = AIService::ExecuteChatWithTools(
			followupMessages,
			settings,
			[](const std::string& toolName, const std::string&, bool& outOk) -> std::string {
				const std::string actualArgs = BuildGeminiToolArgumentsJson(toolName);
				return ExecuteToolCall(toolName, actualArgs, outOk, false);
			});
		report["followup_chat"] = {
			{"ok", followupResult.ok},
			{"cancelled", followupResult.cancelled},
			{"http_status", followupResult.httpStatus},
			{"content", followupResult.content},
			{"reasoning_content_present", !followupResult.reasoningContent.empty()},
			{"reasoning_content_size", followupResult.reasoningContent.size()},
			{"error", followupResult.error},
			{"tool_event_count", followupResult.toolEvents.size()}
		};

		report["ok"] =
			connectionResult.ok &&
			simpleTaskResult.ok &&
			toolChatResult.ok &&
			followupResult.ok &&
			toolChatResult.toolEvents.size() >= 2 &&
			allToolEventsOk;
		return CopyStringToBuffer(DumpJsonPrettySafe(report), buffer, bufferSize);
	}
	catch (const std::exception& ex) {
		nlohmann::json report = {
			{"ok", false},
			{"provider", "gemini"},
			{"model", model},
			{"base_url", (baseUrl != nullptr && baseUrl[0] != '\0') ? baseUrl : defaultBaseUrl},
			{"protocol", "gemini"},
			{"step", step},
			{"error", std::string("exception: ") + ex.what()}
		};
		return CopyStringToBuffer(DumpJsonPrettySafe(report), buffer, bufferSize);
	}
	catch (...) {
		nlohmann::json report = {
			{"ok", false},
			{"provider", "gemini"},
			{"model", model},
			{"base_url", (baseUrl != nullptr && baseUrl[0] != '\0') ? baseUrl : defaultBaseUrl},
			{"protocol", "gemini"},
			{"step", step},
			{"error", "unknown exception"}
		};
		return CopyStringToBuffer(DumpJsonPrettySafe(report), buffer, bufferSize);
	}
}

int RunClaudeIntegrationTestInternal(
	const char* apiKey,
	const char* model,
	const char* baseUrl,
	char* buffer,
	int bufferSize)
{
	if (apiKey == nullptr || model == nullptr) {
		return AUTOLINKER_TEST_STRING_INVALID_ARGUMENT;
	}

	const char* defaultBaseUrl = "https://api.anthropic.com";
	std::string step = "init";
	try {
		AISettings settings = {};
		settings.protocolType = AIProtocolType::Claude;
		settings.thinkingLevel = AIThinkingLevel::Off;
		settings.baseUrl = (baseUrl != nullptr && baseUrl[0] != '\0') ? baseUrl : defaultBaseUrl;
		settings.apiKey = apiKey;
		settings.model = model;
		settings.timeoutMs = 180000;
		settings.temperature = 0;

		nlohmann::json report = BuildClaudeIntegrationResultJson(settings);
		report["step"] = step;

		step = "test_connection";
		report["step"] = step;
		const AIResult connectionResult = AIService::TestConnection(settings);
		report["test_connection"] = {
			{"ok", connectionResult.ok},
			{"http_status", connectionResult.httpStatus},
			{"content", connectionResult.content},
			{"error", connectionResult.error}
		};
		if (!connectionResult.ok) {
			report["ok"] = false;
			return CopyStringToBuffer(DumpJsonPrettySafe(report), buffer, bufferSize);
		}

		step = "simple_task";
		report["step"] = step;
		const AIResult simpleTaskResult = AIService::ExecuteTask(
			AITaskKind::TranslateText,
			"只返回这四个字符：测试通过",
			settings);
		report["simple_task"] = {
			{"ok", simpleTaskResult.ok},
			{"http_status", simpleTaskResult.httpStatus},
			{"content", simpleTaskResult.content},
			{"error", simpleTaskResult.error}
		};
		if (!simpleTaskResult.ok) {
			report["ok"] = false;
			return CopyStringToBuffer(DumpJsonPrettySafe(report), buffer, bufferSize);
		}

		step = "tool_chat";
		report["step"] = step;
		std::vector<AIChatMessage> contextMessages;
		contextMessages.push_back({
			"user",
			"你必须先后调用两个工具：先 fetch_url 读取 https://docs.anthropic.com/en/docs/about-claude/models/overview ，再 extract_web_document 读取 https://docs.anthropic.com/en/docs/agents-and-tools/tool-use/overview 。完成后仅用一行中文回答，格式必须是：模型页已读；工具页已读；工具数=N。",
			"",
			""
		});

		std::vector<std::string> streamedDeltas;
		const AIChatResult toolChatResult = AIService::ExecuteChatWithTools(
			contextMessages,
			settings,
			[](const std::string& toolName, const std::string&, bool& outOk) -> std::string {
				const std::string actualArgs = BuildClaudeToolArgumentsJson(toolName);
				return ExecuteToolCall(toolName, actualArgs, outOk, false);
			},
			[&streamedDeltas](const std::string& deltaText) {
				if (!deltaText.empty()) {
					streamedDeltas.push_back(deltaText);
				}
			});

		nlohmann::json toolEvents = nlohmann::json::array();
		bool allToolEventsOk = true;
		for (const auto& evt : toolChatResult.toolEvents) {
			allToolEventsOk = allToolEventsOk && evt.ok;
			toolEvents.push_back({
				{"name", evt.name},
				{"arguments_json", evt.argumentsJson},
				{"result_json", evt.resultJson},
				{"ok", evt.ok}
			});
		}
		const bool hiddenContextOk = toolChatResult.contextPrefixRawMessagesUtf8.size() >= 2;
		report["tool_chat"] = {
			{"ok", toolChatResult.ok},
			{"cancelled", toolChatResult.cancelled},
			{"http_status", toolChatResult.httpStatus},
			{"content", toolChatResult.content},
			{"reasoning_content_present", !toolChatResult.reasoningContent.empty()},
			{"reasoning_content_size", toolChatResult.reasoningContent.size()},
			{"error", toolChatResult.error},
			{"tool_events", std::move(toolEvents)},
			{"all_tool_events_ok", allToolEventsOk},
			{"hidden_context_ok", hiddenContextOk},
			{"stream_chunk_count", streamedDeltas.size()},
			{"hidden_context_message_count", toolChatResult.contextPrefixRawMessagesUtf8.size()}
		};
		if (!toolChatResult.ok) {
			report["ok"] = false;
			return CopyStringToBuffer(DumpJsonPrettySafe(report), buffer, bufferSize);
		}

		step = "followup_chat";
		report["step"] = step;
		std::vector<AIChatMessage> followupMessages = BuildFollowupMessagesFromChatResult(contextMessages, toolChatResult);
		followupMessages.push_back({
			"assistant",
			toolChatResult.content,
			toolChatResult.reasoningContent,
			""
		});
		followupMessages.push_back({
			"user",
			"只回答：上一轮你实际调用了几个工具？输出阿拉伯数字。",
			"",
			""
		});

		const AIChatResult followupResult = AIService::ExecuteChatWithTools(
			followupMessages,
			settings,
			[](const std::string& toolName, const std::string&, bool& outOk) -> std::string {
				const std::string actualArgs = BuildClaudeToolArgumentsJson(toolName);
				return ExecuteToolCall(toolName, actualArgs, outOk, false);
			});
		report["followup_chat"] = {
			{"ok", followupResult.ok},
			{"cancelled", followupResult.cancelled},
			{"http_status", followupResult.httpStatus},
			{"content", followupResult.content},
			{"reasoning_content_present", !followupResult.reasoningContent.empty()},
			{"reasoning_content_size", followupResult.reasoningContent.size()},
			{"error", followupResult.error},
			{"tool_event_count", followupResult.toolEvents.size()}
		};

		report["ok"] =
			connectionResult.ok &&
			simpleTaskResult.ok &&
			toolChatResult.ok &&
			followupResult.ok &&
			toolChatResult.toolEvents.size() >= 2 &&
			allToolEventsOk &&
			hiddenContextOk;
		return CopyStringToBuffer(DumpJsonPrettySafe(report), buffer, bufferSize);
	}
	catch (const std::exception& ex) {
		nlohmann::json report = {
			{"ok", false},
			{"provider", "claude"},
			{"model", model},
			{"base_url", (baseUrl != nullptr && baseUrl[0] != '\0') ? baseUrl : defaultBaseUrl},
			{"protocol", "claude"},
			{"step", step},
			{"error", std::string("exception: ") + ex.what()}
		};
		return CopyStringToBuffer(DumpJsonPrettySafe(report), buffer, bufferSize);
	}
	catch (...) {
		nlohmann::json report = {
			{"ok", false},
			{"provider", "claude"},
			{"model", model},
			{"base_url", (baseUrl != nullptr && baseUrl[0] != '\0') ? baseUrl : defaultBaseUrl},
			{"protocol", "claude"},
			{"step", step},
			{"error", "unknown exception"}
		};
		return CopyStringToBuffer(DumpJsonPrettySafe(report), buffer, bufferSize);
	}
}

}

extern "C" bool AutoLinkerTest_CompareVersion(const char* left, const char* right, int* outResult)
{
	if (left == nullptr || right == nullptr || outResult == nullptr) {
		return false;
	}

	const Version leftVersion(left);
	const Version rightVersion(right);
	if (leftVersion < rightVersion) {
		*outResult = -1;
	}
	else if (leftVersion > rightVersion) {
		*outResult = 1;
	}
	else {
		*outResult = 0;
	}

	return true;
}

extern "C" int AutoLinkerTest_GetLinkerOutFileName(const char* commandLine, char* buffer, int bufferSize)
{
	if (commandLine == nullptr) {
		return AUTOLINKER_TEST_STRING_INVALID_ARGUMENT;
	}

	return CopyStringToBuffer(GetLinkerCommandOutFileName(commandLine), buffer, bufferSize);
}

extern "C" int AutoLinkerTest_GetLinkerKrnlnFileName(const char* commandLine, char* buffer, int bufferSize)
{
	if (commandLine == nullptr) {
		return AUTOLINKER_TEST_STRING_INVALID_ARGUMENT;
	}

	return CopyStringToBuffer(GetLinkerCommandKrnlnFileName(commandLine), buffer, bufferSize);
}

extern "C" int AutoLinkerTest_ExtractBetweenDashes(const char* text, char* buffer, int bufferSize)
{
	if (text == nullptr) {
		return AUTOLINKER_TEST_STRING_INVALID_ARGUMENT;
	}

	return CopyStringToBuffer(ExtractBetweenDashes(text), buffer, bufferSize);
}

extern "C" int AutoLinkerTest_GetVersionText(char* buffer, int bufferSize)
{
	return CopyStringToBuffer(AUTOLINKER_VERSION, buffer, bufferSize);
}

extern "C" int AutoLinkerTest_RunGameAnalyticsSelfTest(char* buffer, int bufferSize)
{
	return CopyStringToBuffer(GameAnalyticsClient::BuildSelfTestReportJson(), buffer, bufferSize);
}

extern "C" int AutoLinkerTest_RunDeepSeekModelIntegrationTest(
	const char* apiKey,
	const char* model,
	const char* baseUrl,
	char* buffer,
	int bufferSize)
{
	if (apiKey == nullptr || model == nullptr) {
		return AUTOLINKER_TEST_STRING_INVALID_ARGUMENT;
	}

	std::string step = "init";
	try {
		AISettings settings = {};
		settings.protocolType = AIProtocolType::OpenAI;
		settings.thinkingLevel = AIThinkingLevel::High;
		settings.baseUrl = (baseUrl != nullptr && baseUrl[0] != '\0') ? baseUrl : "https://api.deepseek.com";
		settings.apiKey = apiKey;
		settings.model = model;
		settings.timeoutMs = 180000;
		settings.temperature = 0;

		nlohmann::json report = BuildDeepSeekIntegrationResultJson(settings);
		report["step"] = step;

		step = "test_connection";
		report["step"] = step;
		const AIResult connectionResult = AIService::TestConnection(settings);
		report["test_connection"] = {
			{"ok", connectionResult.ok},
			{"http_status", connectionResult.httpStatus},
			{"content", connectionResult.content},
			{"error", connectionResult.error}
		};
		if (!connectionResult.ok) {
			report["ok"] = false;
			return CopyStringToBuffer(DumpJsonPrettySafe(report), buffer, bufferSize);
		}

		step = "simple_task";
		report["step"] = step;
		const AIResult simpleTaskResult = AIService::ExecuteTask(
			AITaskKind::TranslateText,
			"只返回这四个字符：测试通过",
			settings);
		report["simple_task"] = {
			{"ok", simpleTaskResult.ok},
			{"http_status", simpleTaskResult.httpStatus},
			{"content", simpleTaskResult.content},
			{"error", simpleTaskResult.error}
		};
		if (!simpleTaskResult.ok) {
			report["ok"] = false;
			return CopyStringToBuffer(DumpJsonPrettySafe(report), buffer, bufferSize);
		}

		step = "tool_chat";
		report["step"] = step;
		std::vector<AIChatMessage> contextMessages;
		contextMessages.push_back({
			"user",
			"你必须先后调用两个工具：先 fetch_url 读取 https://api-docs.deepseek.com/quick_start/rate_limit ，再 extract_web_document 读取 https://api-docs.deepseek.com/guides/thinking_mode 。完成后仅用一行中文回答，格式必须是：限速页已读；思考页已读；工具数=N。",
			"",
			""
		});

		std::vector<std::string> streamedDeltas;
		const AIChatResult toolChatResult = AIService::ExecuteChatWithTools(
			contextMessages,
			settings,
			[](const std::string& toolName, const std::string&, bool& outOk) -> std::string {
				const std::string actualArgs = BuildDeepSeekToolArgumentsJson(toolName);
				return ExecuteToolCall(toolName, actualArgs, outOk, false);
			},
			[&streamedDeltas](const std::string& deltaText) {
				if (!deltaText.empty()) {
					streamedDeltas.push_back(deltaText);
				}
			});

		nlohmann::json toolEvents = nlohmann::json::array();
		for (const auto& evt : toolChatResult.toolEvents) {
			toolEvents.push_back({
				{"name", evt.name},
				{"arguments_json", evt.argumentsJson},
				{"result_json", evt.resultJson},
				{"ok", evt.ok}
			});
		}
		report["tool_chat"] = {
			{"ok", toolChatResult.ok},
			{"cancelled", toolChatResult.cancelled},
			{"http_status", toolChatResult.httpStatus},
			{"content", toolChatResult.content},
			{"reasoning_content_present", !toolChatResult.reasoningContent.empty()},
			{"reasoning_content_size", toolChatResult.reasoningContent.size()},
			{"error", toolChatResult.error},
			{"tool_events", std::move(toolEvents)},
			{"stream_chunk_count", streamedDeltas.size()},
			{"hidden_context_message_count", toolChatResult.contextPrefixRawMessagesUtf8.size()}
		};
		if (!toolChatResult.ok) {
			report["ok"] = false;
			return CopyStringToBuffer(DumpJsonPrettySafe(report), buffer, bufferSize);
		}

		step = "followup_chat";
		report["step"] = step;
		std::vector<AIChatMessage> followupMessages;
		followupMessages.push_back(contextMessages.front());
		for (const auto& rawMessageJsonUtf8 : toolChatResult.contextPrefixRawMessagesUtf8) {
			nlohmann::json parsed;
			try {
				parsed = nlohmann::json::parse(rawMessageJsonUtf8);
			}
			catch (...) {
				continue;
			}
			if (!parsed.is_object()) {
				continue;
			}

			const std::string role = parsed.value("role", std::string());
			if (role == "assistant") {
				followupMessages.push_back({
					"assistant",
					ExtractMessageContentText(parsed),
					parsed.value("reasoning_content", std::string()),
					rawMessageJsonUtf8
				});
			}
			else if (role == "tool") {
				followupMessages.push_back({
					"tool",
					ExtractMessageContentText(parsed),
					"",
					rawMessageJsonUtf8
				});
			}
		}
		followupMessages.push_back({
			"assistant",
			toolChatResult.content,
			toolChatResult.reasoningContent,
			""
		});
		followupMessages.push_back({
			"user",
			"只回答：上一轮你实际调用了几个工具？输出阿拉伯数字。",
			"",
			""
		});

		const AIChatResult followupResult = AIService::ExecuteChatWithTools(
			followupMessages,
			settings,
			[](const std::string& toolName, const std::string&, bool& outOk) -> std::string {
				const std::string actualArgs = BuildDeepSeekToolArgumentsJson(toolName);
				return ExecuteToolCall(toolName, actualArgs, outOk, false);
			});
		report["followup_chat"] = {
			{"ok", followupResult.ok},
			{"cancelled", followupResult.cancelled},
			{"http_status", followupResult.httpStatus},
			{"content", followupResult.content},
			{"reasoning_content_present", !followupResult.reasoningContent.empty()},
			{"reasoning_content_size", followupResult.reasoningContent.size()},
			{"error", followupResult.error},
			{"tool_event_count", followupResult.toolEvents.size()}
		};

		report["ok"] =
			connectionResult.ok &&
			simpleTaskResult.ok &&
			toolChatResult.ok &&
			followupResult.ok &&
			toolChatResult.toolEvents.size() >= 2;
		return CopyStringToBuffer(DumpJsonPrettySafe(report), buffer, bufferSize);
	}
	catch (const std::exception& ex) {
		nlohmann::json report = {
			{"ok", false},
			{"provider", "deepseek"},
			{"model", model},
			{"base_url", (baseUrl != nullptr && baseUrl[0] != '\0') ? baseUrl : "https://api.deepseek.com"},
			{"step", step},
			{"error", std::string("exception: ") + ex.what()}
		};
		return CopyStringToBuffer(DumpJsonPrettySafe(report), buffer, bufferSize);
	}
	catch (...) {
		nlohmann::json report = {
			{"ok", false},
			{"provider", "deepseek"},
			{"model", model},
			{"base_url", (baseUrl != nullptr && baseUrl[0] != '\0') ? baseUrl : "https://api.deepseek.com"},
			{"step", step},
			{"error", "unknown exception"}
		};
		return CopyStringToBuffer(DumpJsonPrettySafe(report), buffer, bufferSize);
	}
}

extern "C" int AutoLinkerTest_RunDeepSeekConnectionOnly(
	const char* apiKey,
	const char* model,
	const char* baseUrl,
	char* buffer,
	int bufferSize)
{
	if (apiKey == nullptr || model == nullptr) {
		return AUTOLINKER_TEST_STRING_INVALID_ARGUMENT;
	}

	AISettings settings = {};
	settings.protocolType = AIProtocolType::OpenAI;
	settings.thinkingLevel = AIThinkingLevel::High;
	settings.baseUrl = (baseUrl != nullptr && baseUrl[0] != '\0') ? baseUrl : "https://api.deepseek.com";
	settings.apiKey = apiKey;
	settings.model = model;
	settings.timeoutMs = 180000;
	settings.temperature = 0;

	const AIResult connectionResult = AIService::TestConnection(settings);
	nlohmann::json report = BuildDeepSeekIntegrationResultJson(settings);
	report["ok"] = connectionResult.ok;
	report["test_connection"] = {
		{"ok", connectionResult.ok},
		{"http_status", connectionResult.httpStatus},
		{"content", connectionResult.content},
		{"error", connectionResult.error}
	};
	return CopyStringToBuffer(DumpJsonPrettySafe(report), buffer, bufferSize);
}

extern "C" int AutoLinkerTest_RunOpenAIChatIntegrationTest(
	const char* apiKey,
	const char* model,
	const char* baseUrl,
	char* buffer,
	int bufferSize)
{
	return RunOpenAIIntegrationTestInternal(
		AIProtocolType::OpenAI,
		apiKey,
		model,
		baseUrl,
		buffer,
		bufferSize);
}

extern "C" int AutoLinkerTest_RunOpenAIResponsesIntegrationTest(
	const char* apiKey,
	const char* model,
	const char* baseUrl,
	char* buffer,
	int bufferSize)
{
	return RunOpenAIIntegrationTestInternal(
		AIProtocolType::OpenAIResponses,
		apiKey,
		model,
		baseUrl,
		buffer,
		bufferSize);
}

extern "C" int AutoLinkerTest_RunGeminiIntegrationTest(
	const char* apiKey,
	const char* model,
	const char* baseUrl,
	char* buffer,
	int bufferSize)
{
	return RunGeminiIntegrationTestInternal(
		apiKey,
		model,
		baseUrl,
		buffer,
		bufferSize);
}

extern "C" int AutoLinkerTest_RunClaudeIntegrationTest(
	const char* apiKey,
	const char* model,
	const char* baseUrl,
	char* buffer,
	int bufferSize)
{
	return RunClaudeIntegrationTestInternal(
		apiKey,
		model,
		baseUrl,
		buffer,
		bufferSize);
}
