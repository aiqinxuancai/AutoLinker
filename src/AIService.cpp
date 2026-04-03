#include "AIService.h"

#include <algorithm>
#include <cctype>
#include <filesystem>
#include <format>
#include <string_view>
#include <Windows.h>

#include "..\\thirdparty\\json.hpp"

#include "ConfigManager.h"
#include "Global.h"
#include "IDEFacade.h"
#include "WinINetUtil.h"
#include <chrono>

namespace {
using PerfClock = std::chrono::steady_clock;

long long ElapsedMs(const PerfClock::time_point& start)
{
	return static_cast<long long>(
		std::chrono::duration_cast<std::chrono::milliseconds>(PerfClock::now() - start).count());
}

std::string ToLowerAsciiCopy(const std::string& text)
{
	std::string lowered = text;
	std::transform(lowered.begin(), lowered.end(), lowered.begin(), [](unsigned char c) {
		return static_cast<char>(std::tolower(c));
	});
	return lowered;
}

bool EndsWithInsensitive(const std::string& text, const std::string& suffix)
{
	if (text.size() < suffix.size()) {
		return false;
	}
	const std::string tail = text.substr(text.size() - suffix.size());
	return ToLowerAsciiCopy(tail) == ToLowerAsciiCopy(suffix);
}

std::string DetectProjectTypeText()
{
	struct Candidate {
		INT fnCode;
		const char* text;
	};

	const Candidate candidates[] = {
		{FN_COMPILE_WINDOWS_DLL, "DLL"},
		{FN_COMPILE_WINDOWS_EXE, "窗口程序 EXE"},
		{FN_COMPILE_WINDOWS_CONOLE_EXE, "控制台程序 EXE"},
		{FN_COMPILE_WINDOWS_ECOM, "易模块"}
	};

	auto& ide = IDEFacade::Instance();
	std::string detected;
	for (const auto& candidate : candidates) {
		if (!ide.IsFunctionEnabled(candidate.fnCode)) {
			continue;
		}
		if (!detected.empty()) {
			detected += " / ";
		}
		detected += candidate.text;
	}

	return detected.empty() ? std::string("未知") : detected;
}

std::string TruncateForLog(const std::string& text, size_t maxLen = 240)
{
	if (text.size() <= maxLen) {
		return text;
	}
	return text.substr(0, maxLen) + "...";
}

constexpr int kAiRequestRetryCount = 5;

bool IsSuccessfulHttpStatus(int statusCode)
{
	return statusCode >= 200 && statusCode < 300;
}

bool IsRetryableHttpStatus(int statusCode)
{
	return statusCode == 0 ||
		statusCode == 408 ||
		statusCode == 409 ||
		statusCode == 425 ||
		statusCode == 429 ||
		statusCode == 500 ||
		statusCode == 502 ||
		statusCode == 503 ||
		statusCode == 504;
}

bool ContainsRetryableTransportHint(const std::string& responseBody)
{
	const std::string lower = ToLowerAsciiCopy(responseBody);
	return lower.find("error in internetopen") != std::string::npos ||
		lower.find("error in internetcrackurl") != std::string::npos ||
		lower.find("error in internetconnect") != std::string::npos ||
		lower.find("error in httpopenrequest") != std::string::npos ||
		lower.find("error in httpsendrequest") != std::string::npos ||
		lower.find("timeout") != std::string::npos ||
		lower.find("timed out") != std::string::npos ||
		lower.find("cannot connect") != std::string::npos ||
		lower.find("connection") != std::string::npos;
}

bool ShouldRetryAiHttpRequest(int statusCode, const std::string& responseBody)
{
	return IsRetryableHttpStatus(statusCode) || ContainsRetryableTransportHint(responseBody);
}

DWORD ComputeAiRetryDelayMs(int retryIndex)
{
	switch (retryIndex) {
	case 0:
		return 250;
	case 1:
		return 500;
	case 2:
		return 1000;
	case 3:
		return 1500;
	default:
		return 2000;
	}
}

void LogAiRetryAttempt(const std::string& tag, int nextAttemptIndex, int statusCode, const std::string& responseBody)
{
	OutputStringToELog(std::format(
		"[AI Chat][Retry] {} attempt {}/{} http={} reason={}",
		tag,
		nextAttemptIndex,
		kAiRequestRetryCount + 1,
		statusCode,
		TruncateForLog(responseBody, 120)));
}

std::pair<std::string, int> PerformPostRequestWithRetry(
	const std::string& url,
	const std::string& postData,
	const std::string& customHeaders,
	int timeout,
	bool autoCookies,
	bool neverRedirect,
	const char* retryTag)
{
	std::pair<std::string, int> lastResult;
	for (int attempt = 0; attempt <= kAiRequestRetryCount; ++attempt) {
		lastResult = PerformPostRequest(url, postData, customHeaders, timeout, autoCookies, neverRedirect);
		if (!ShouldRetryAiHttpRequest(lastResult.second, lastResult.first) || attempt >= kAiRequestRetryCount) {
			return lastResult;
		}

		LogAiRetryAttempt(retryTag == nullptr ? "post" : retryTag, attempt + 2, lastResult.second, lastResult.first);
		::Sleep(ComputeAiRetryDelayMs(attempt));
	}
	return lastResult;
}

std::pair<std::string, int> PerformPostRequestStreamingWithRetry(
	const std::string& url,
	const std::string& postData,
	const std::function<bool(const std::string& chunk)>& onChunk,
	const std::string& customHeaders,
	int timeout,
	bool autoCookies,
	bool neverRedirect,
	const char* retryTag)
{
	std::pair<std::string, int> lastResult;
	for (int attempt = 0; attempt <= kAiRequestRetryCount; ++attempt) {
		bool sawChunk = false;
		lastResult = PerformPostRequestStreaming(
			url,
			postData,
			[&onChunk, &sawChunk](const std::string& chunk) -> bool {
				if (!chunk.empty()) {
					sawChunk = true;
				}
				return onChunk ? onChunk(chunk) : true;
			},
			customHeaders,
			timeout,
			autoCookies,
			neverRedirect);
		const bool streamAccepted = sawChunk && IsSuccessfulHttpStatus(lastResult.second);
		if (streamAccepted || !ShouldRetryAiHttpRequest(lastResult.second, lastResult.first) || attempt >= kAiRequestRetryCount) {
			return lastResult;
		}

		LogAiRetryAttempt(retryTag == nullptr ? "stream" : retryTag, attempt + 2, lastResult.second, lastResult.first);
		::Sleep(ComputeAiRetryDelayMs(attempt));
	}
	return lastResult;
}

std::string LocalToUtf8(const std::string& text);

size_t ClampUtf8PrefixBoundary(const std::string& text, size_t maxBytes)
{
	size_t end = (std::min)(maxBytes, text.size());
	while (end > 0 && end < text.size() &&
		(static_cast<unsigned char>(text[end]) & 0xC0) == 0x80) {
		--end;
	}
	return end;
}

size_t ClampUtf8SuffixStartBoundary(const std::string& text, size_t tailBytes)
{
	if (tailBytes >= text.size()) {
		return 0;
	}

	size_t start = text.size() - tailBytes;
	while (start < text.size() &&
		(static_cast<unsigned char>(text[start]) & 0xC0) == 0x80) {
		++start;
	}
	return start;
}

std::string TruncateUtf8Text(const std::string& text, size_t maxBytes)
{
	if (text.size() <= maxBytes) {
		return text;
	}

	const size_t end = ClampUtf8PrefixBoundary(text, maxBytes);
	return text.substr(0, end) + std::format("...[truncated {} bytes]", text.size() - end);
}

std::string BuildUtf8Excerpt(const std::string& text, size_t headBytes, size_t tailBytes)
{
	if (text.size() <= headBytes + tailBytes + 64) {
		return text;
	}

	const size_t headEnd = ClampUtf8PrefixBoundary(text, headBytes);
	const size_t tailStart = ClampUtf8SuffixStartBoundary(text, tailBytes);
	if (tailStart <= headEnd) {
		return TruncateUtf8Text(text, headBytes + tailBytes);
	}

	return text.substr(0, headEnd) +
		std::format("\n...[truncated {} bytes]...\n", tailStart - headEnd) +
		text.substr(tailStart);
}

bool EndsWithAsciiInsensitive(std::string_view text, std::string_view suffix)
{
	if (text.size() < suffix.size()) {
		return false;
	}

	const size_t offset = text.size() - suffix.size();
	for (size_t i = 0; i < suffix.size(); ++i) {
		const unsigned char left = static_cast<unsigned char>(text[offset + i]);
		const unsigned char right = static_cast<unsigned char>(suffix[i]);
		if (std::tolower(left) != std::tolower(right)) {
			return false;
		}
	}
	return true;
}

std::string ToLowerAsciiCopy(std::string_view text)
{
	std::string lowered(text.begin(), text.end());
	std::transform(lowered.begin(), lowered.end(), lowered.begin(), [](unsigned char c) {
		return static_cast<char>(std::tolower(c));
	});
	return lowered;
}

bool IsCodeLikeKey(std::string_view key)
{
	const std::string lowered = ToLowerAsciiCopy(key);
	return lowered == "code" ||
		lowered == "proposed_code" ||
		lowered == "plain_text" ||
		EndsWithAsciiInsensitive(lowered, "_code");
}

bool IsTraceLikeKey(std::string_view key)
{
	const std::string lowered = ToLowerAsciiCopy(key);
	return lowered == "trace" ||
		EndsWithAsciiInsensitive(lowered, "_trace");
}

size_t GetCompactArrayLimit(std::string_view key)
{
	const std::string lowered = ToLowerAsciiCopy(key);
	if (lowered == "hunks") {
		return 4;
	}
	if (lowered == "matches") {
		return 5;
	}
	if (lowered == "symbols") {
		return 16;
	}
	if (lowered == "results") {
		return 8;
	}
	return 6;
}

nlohmann::json CompactToolContextJsonValue(const nlohmann::json& value, std::string_view key, int depth)
{
	if (depth >= 6) {
		return "[omitted: max depth]";
	}

	if (value.is_null() || value.is_boolean() || value.is_number()) {
		return value;
	}

	if (value.is_string()) {
		const std::string text = value.get<std::string>();
		if (IsCodeLikeKey(key)) {
			return BuildUtf8Excerpt(text, 2400, 900);
		}
		if (IsTraceLikeKey(key)) {
			return TruncateUtf8Text(text, 1200);
		}
		return TruncateUtf8Text(text, 800);
	}

	if (value.is_array()) {
		const size_t limit = GetCompactArrayLimit(key);
		nlohmann::json out = nlohmann::json::array();
		for (size_t i = 0; i < value.size() && i < limit; ++i) {
			out.push_back(CompactToolContextJsonValue(value[i], key, depth + 1));
		}
		if (value.size() > limit) {
			out.push_back({
				{"_truncated", true},
				{"omitted_items", value.size() - limit}
			});
		}
		return out;
	}

	if (value.is_object()) {
		nlohmann::json out = nlohmann::json::object();
		for (auto it = value.begin(); it != value.end(); ++it) {
			out[it.key()] = CompactToolContextJsonValue(it.value(), it.key(), depth + 1);
		}
		return out;
	}

	return TruncateUtf8Text(value.dump(), 800);
}

struct CompactToolResultPayload {
	std::string textUtf8;
	nlohmann::json jsonValue = nlohmann::json::object();
};

CompactToolResultPayload BuildCompactToolResultPayload(const std::string& toolName, const std::string& toolResultLocal)
{
	CompactToolResultPayload payload;
	const std::string resultUtf8 = LocalToUtf8(toolResultLocal);

	try {
		nlohmann::json parsed = nlohmann::json::parse(resultUtf8);
		nlohmann::json compact = CompactToolContextJsonValue(parsed, "", 0);
		if (!compact.is_object()) {
			payload.jsonValue = {
				{"tool_name", toolName},
				{"result", compact}
			};
		}
		else {
			payload.jsonValue = std::move(compact);
		}
	}
	catch (...) {
		payload.jsonValue = {
			{"tool_name", toolName},
			{"ok", false},
			{"text", BuildUtf8Excerpt(resultUtf8, 1800, 600)}
		};
	}

	payload.textUtf8 = payload.jsonValue.dump();
	return payload;
}

int GetMaxToolRounds(const AISettings& settings)
{
	return (std::clamp)(settings.maxToolRounds, 4, 64);
}

std::string BuildToolRoundsExceededError(int maxToolRounds, const std::vector<AIChatToolEvent>& toolEvents)
{
	std::string message = std::format("tool call rounds exceeded limit ({})", maxToolRounds);
	if (!toolEvents.empty()) {
		message += " after ";
		message += std::to_string(toolEvents.size());
		message += " tool calls";
	}
	return message;
}

bool IsValidUtf8(const std::string& text)
{
	if (text.empty()) {
		return true;
	}
	return MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS, text.data(), static_cast<int>(text.size()), nullptr, 0) > 0;
}

std::string ConvertCodePage(const std::string& text, UINT fromCodePage, UINT toCodePage, DWORD fromFlags = 0)
{
	if (text.empty()) {
		return std::string();
	}

	const int wideLen = MultiByteToWideChar(
		fromCodePage,
		fromFlags,
		text.data(),
		static_cast<int>(text.size()),
		nullptr,
		0);
	if (wideLen <= 0) {
		return text;
	}

	std::wstring wide(static_cast<size_t>(wideLen), L'\0');
	if (MultiByteToWideChar(
		fromCodePage,
		fromFlags,
		text.data(),
		static_cast<int>(text.size()),
		&wide[0],
		wideLen) <= 0) {
		return text;
	}

	const int outLen = WideCharToMultiByte(
		toCodePage,
		0,
		wide.data(),
		wideLen,
		nullptr,
		0,
		nullptr,
		nullptr);
	if (outLen <= 0) {
		return text;
	}

	std::string out(static_cast<size_t>(outLen), '\0');
	if (WideCharToMultiByte(
		toCodePage,
		0,
		wide.data(),
		wideLen,
		&out[0],
		outLen,
		nullptr,
		nullptr) <= 0) {
		return text;
	}
	return out;
}

std::string LocalToUtf8(const std::string& text)
{
	// AutoLinker/IDE strings are typically local ANSI (GBK on zh-CN Windows).
	// Convert before feeding nlohmann::json, which requires UTF-8.
	if (text.empty()) {
		return std::string();
	}
	if (IsValidUtf8(text)) {
		return text;
	}
	return ConvertCodePage(text, CP_ACP, CP_UTF8, 0);
}

std::string Utf8ToLocal(const std::string& text)
{
	if (text.empty()) {
		return std::string();
	}
	if (!IsValidUtf8(text)) {
		return text;
	}
	return ConvertCodePage(text, CP_UTF8, CP_ACP, MB_ERR_INVALID_CHARS);
}

std::string RemoveCodeFence(const std::string& text)
{
	const std::string content = AIService::Trim(text);
	size_t fenceBegin = content.find("```");
	if (fenceBegin == std::string::npos) {
		return content;
	}

	size_t firstLineEnd = content.find('\n', fenceBegin + 3);
	if (firstLineEnd == std::string::npos) {
		return content;
	}

	size_t fenceEnd = content.find("```", firstLineEnd + 1);
	if (fenceEnd == std::string::npos) {
		return content;
	}

	return AIService::Trim(content.substr(firstLineEnd + 1, fenceEnd - (firstLineEnd + 1)));
}

std::string MergeMessageContentUtf8(const nlohmann::json& message)
{
	std::string merged;
	if (!message.contains("content")) {
		return merged;
	}

	const nlohmann::json& content = message["content"];
	if (content.is_string()) {
		return content.get<std::string>();
	}
	if (!content.is_array()) {
		return merged;
	}

	for (const auto& item : content) {
		if (item.is_string()) {
			merged += item.get<std::string>();
			continue;
		}
		if (item.is_object() && item.contains("text") && item["text"].is_string()) {
			merged += item["text"].get<std::string>();
		}
	}
	return merged;
}

bool ExtractChatResponseMessage(const nlohmann::json& parsed, nlohmann::json& outMessage, std::string& outError)
{
	if (!parsed.contains("choices") || !parsed["choices"].is_array() || parsed["choices"].empty()) {
		outError = "AI response choices is empty";
		return false;
	}
	const nlohmann::json& choice = parsed["choices"][0];
	if (!choice.contains("message") || !choice["message"].is_object()) {
		outError = "AI response message missing";
		return false;
	}
	outMessage = choice["message"];
	return true;
}

struct StreamToolCallState {
	std::string id;
	std::string name;
	std::string arguments;
};

struct ChatStreamParseState {
	bool sawDataEvent = false;
	std::string pendingLine;
	std::string mergedUtf8;
	std::vector<StreamToolCallState> toolCalls;
	std::string parseError;
};

StreamToolCallState& EnsureToolCallSlot(std::vector<StreamToolCallState>& toolCalls, size_t index)
{
	if (toolCalls.size() <= index) {
		toolCalls.resize(index + 1);
	}
	return toolCalls[index];
}

bool ProcessStreamDataPayload(
	const std::string& payload,
	ChatStreamParseState& state,
	const std::function<void(const std::string& deltaText)>& streamCallback)
{
	if (payload.empty()) {
		return true;
	}
	if (payload == "[DONE]") {
		state.sawDataEvent = true;
		return true;
	}

	state.sawDataEvent = true;
	nlohmann::json packet;
	try {
		packet = nlohmann::json::parse(payload);
	}
	catch (const std::exception& ex) {
		state.parseError = std::string("Failed to parse streaming chunk JSON: ") + ex.what();
		return false;
	}

	if (packet.contains("error") && packet["error"].is_object()) {
		const auto& err = packet["error"];
		if (err.contains("message") && err["message"].is_string()) {
			state.parseError = Utf8ToLocal(err["message"].get<std::string>());
		}
		else {
			state.parseError = "AI streaming response contains error";
		}
		return false;
	}

	if (!packet.contains("choices") || !packet["choices"].is_array() || packet["choices"].empty()) {
		return true;
	}

	const auto& choice = packet["choices"][0];
	if (!choice.contains("delta") || !choice["delta"].is_object()) {
		return true;
	}
	const auto& delta = choice["delta"];

	const std::string deltaContentUtf8 = MergeMessageContentUtf8(delta);
	if (!deltaContentUtf8.empty()) {
		state.mergedUtf8 += deltaContentUtf8;
		if (streamCallback) {
			streamCallback(Utf8ToLocal(deltaContentUtf8));
		}
	}

	if (!delta.contains("tool_calls") || !delta["tool_calls"].is_array()) {
		return true;
	}

	for (const auto& toolCallDelta : delta["tool_calls"]) {
		if (!toolCallDelta.is_object()) {
			continue;
		}

		size_t index = state.toolCalls.size();
		if (toolCallDelta.contains("index") && toolCallDelta["index"].is_number_integer()) {
			const int idx = toolCallDelta["index"].get<int>();
			if (idx >= 0) {
				index = static_cast<size_t>(idx);
			}
		}

		auto& slot = EnsureToolCallSlot(state.toolCalls, index);
		if (toolCallDelta.contains("id") && toolCallDelta["id"].is_string()) {
			const std::string deltaId = toolCallDelta["id"].get<std::string>();
			if (!deltaId.empty()) {
				slot.id = deltaId;
			}
		}

		if (!toolCallDelta.contains("function") || !toolCallDelta["function"].is_object()) {
			continue;
		}

		const auto& fn = toolCallDelta["function"];
		if (fn.contains("name") && fn["name"].is_string()) {
			slot.name += fn["name"].get<std::string>();
		}
		if (fn.contains("arguments") && fn["arguments"].is_string()) {
			slot.arguments += fn["arguments"].get<std::string>();
		}
	}
	return true;
}

bool ProcessStreamLine(
	const std::string& rawLine,
	ChatStreamParseState& state,
	const std::function<void(const std::string& deltaText)>& streamCallback)
{
	std::string line = rawLine;
	if (!line.empty() && line.back() == '\r') {
		line.pop_back();
	}
	if (line.empty()) {
		return true;
	}
	if (line.rfind("data:", 0) != 0) {
		return true;
	}

	std::string payload = line.substr(5);
	if (!payload.empty() && payload[0] == ' ') {
		payload.erase(payload.begin());
	}
	return ProcessStreamDataPayload(payload, state, streamCallback);
}

bool ConsumeStreamChunk(
	const std::string& chunk,
	ChatStreamParseState& state,
	const std::function<void(const std::string& deltaText)>& streamCallback)
{
	state.pendingLine += chunk;
	size_t lineEnd = 0;
	while ((lineEnd = state.pendingLine.find('\n')) != std::string::npos) {
		const std::string line = state.pendingLine.substr(0, lineEnd);
		state.pendingLine.erase(0, lineEnd + 1);
		if (!ProcessStreamLine(line, state, streamCallback)) {
			return false;
		}
	}
	return true;
}

bool FlushStreamParseState(
	ChatStreamParseState& state,
	const std::function<void(const std::string& deltaText)>& streamCallback)
{
	if (state.pendingLine.empty()) {
		return true;
	}
	const std::string line = state.pendingLine;
	state.pendingLine.clear();
	return ProcessStreamLine(line, state, streamCallback);
}

nlohmann::json BuildAssistantMessageFromStreamState(const ChatStreamParseState& state)
{
	nlohmann::json message;
	message["role"] = "assistant";

	if (!state.toolCalls.empty()) {
		message["content"] = state.mergedUtf8;
		message["tool_calls"] = nlohmann::json::array();
		for (size_t i = 0; i < state.toolCalls.size(); ++i) {
			const auto& call = state.toolCalls[i];
			std::string callId = call.id;
			if (callId.empty()) {
				callId = std::format("call_auto_{}", i + 1);
			}
			message["tool_calls"].push_back({
				{"id", callId},
				{"type", "function"},
				{"function", {
					{"name", call.name},
					{"arguments", call.arguments}
				}}
			});
		}
		return message;
	}

	message["content"] = state.mergedUtf8;
	return message;
}

nlohmann::json BuildPublicToolCatalog()
{
	nlohmann::json tools = nlohmann::json::array();
	tools.push_back({
		{"name", "get_current_page_code"},
		{"description", "Get complete source code of the current IDE page, together with current page name and page type."},
		{"inputSchema", {
			{"type", "object"},
			{"properties", nlohmann::json::object()},
			{"additionalProperties", false}
		}}
	});
	tools.push_back({
		{"name", "get_current_page_info"},
		{"description", "Get current IDE page name, page type and the trace/source used to resolve that page name."},
		{"inputSchema", {
			{"type", "object"},
			{"properties", nlohmann::json::object()},
			{"additionalProperties", false}
		}}
	});
	tools.push_back({
		{"name", "get_current_eide_info"},
		{"description", "Get current E-language IDE instance information, including current source file path, current page info, MCP port/endpoint, process id and executable path."},
		{"inputSchema", {
			{"type", "object"},
			{"properties", nlohmann::json::object()},
			{"additionalProperties", false}
		}}
	});
	tools.push_back({
		{"name", "list_local_mcp_instances"},
		{"description", "List other AutoLinker local MCP instances currently running on this machine, including instance_id, pid, port, endpoint and cached IDE hints."},
		{"inputSchema", {
			{"type", "object"},
			{"properties", nlohmann::json::object()},
			{"additionalProperties", false}
		}}
	});
	tools.push_back({
		{"name", "call_local_mcp_instance_tool"},
		{"description", "Forward one MCP tool call to another local AutoLinker instance discovered by list_local_mcp_instances. This is the generic cross-instance method; it is stateless and avoids global switching side effects."},
		{"inputSchema", {
			{"type", "object"},
			{"properties", {
				{"instance_id", {{"type", "string"}, {"description", "Target instance_id returned by list_local_mcp_instances."}}},
				{"port", {{"type", "integer"}, {"minimum", 1}, {"maximum", 65535}, {"description", "Optional target port when instance_id is unknown."}}},
				{"tool_name", {{"type", "string"}}},
				{"arguments", {{"type", "object"}, {"description", "Arguments object passed through to the target tool."}}},
				{"timeout_seconds", {{"type", "integer"}, {"minimum", 1}, {"maximum", 120}}}
			}},
			{"required", nlohmann::json::array({"tool_name"})},
			{"additionalProperties", false}
		}}
	});
	tools.push_back({
		{"name", "list_imported_modules"},
		{"description", "List imported ECOM/e-module paths from the current project. This only lists modules; actual public declarations must be queried separately."},
		{"inputSchema", {
			{"type", "object"},
			{"properties", nlohmann::json::object()},
			{"additionalProperties", false}
		}}
	});
	tools.push_back({
		{"name", "list_support_libraries"},
		{"description", "List support libraries currently selected by the IDE. Returns basic parsed fields and the raw support-library info text returned by the IDE."},
		{"inputSchema", {
			{"type", "object"},
			{"properties", nlohmann::json::object()},
			{"additionalProperties", false}
		}}
	});
	tools.push_back({
		{"name", "get_support_library_info"},
		{"description", "Get one support library's public info. Prefer file_path when available: the tool will parse GetNewInf/lib2.h structures and return commands, constants and data types. If no file path can be resolved, it falls back to the IDE support-library info text."},
		{"inputSchema", {
			{"type", "object"},
			{"properties", {
				{"index", {{"type", "integer"}, {"minimum", 0}}},
				{"name", {{"type", "string"}, {"description", "Support library display name or file name."}}},
				{"file_path", {{"type", "string"}, {"description", "Full support library file path, preferred when known."}}}
			}},
			{"additionalProperties", false}
		}}
	});
	tools.push_back({
		{"name", "search_public_code"},
		{"description", "Complete unified search across current IDE project source hits, module public declarations, and support-library public declarations. Supports one keyword, multiple keywords, or regex. Results include target_type and read_tool so the caller can continue with read_project_search_result_code, read_module_public_code, or read_support_library_public_code."},
		{"inputSchema", {
			{"type", "object"},
			{"properties", {
				{"keyword", {{"type", "string"}, {"description", "Single keyword substring search."}}},
				{"keywords", {{"type", "array"}, {"items", {{"type", "string"}}}, {"description", "Multiple keyword substrings. Default keyword_mode is all."}}},
				{"keyword_mode", {{"type", "string"}, {"enum", nlohmann::json::array({"all", "any"})}}}, 
				{"regex", {{"type", "string"}, {"description", "Optional ECMAScript regex tested line-by-line."}}},
				{"regex_flags", {{"type", "string"}, {"description", "Optional regex flags. Currently supports i for ignore-case."}}},
				{"case_sensitive", {{"type", "boolean"}}},
				{"target_types", {{"type", "array"}, {"items", {{"type", "string"}, {"enum", nlohmann::json::array({"project", "module", "support_library"})}}}}},
				{"module_name", {{"type", "string"}}},
				{"module_path", {{"type", "string"}}},
				{"support_library_index", {{"type", "integer"}, {"minimum", 0}}}, 
				{"support_library_name", {{"type", "string"}}},
				{"support_library_path", {{"type", "string"}}},
				{"limit", {{"type", "integer"}, {"minimum", 1}, {"maximum", 500}}}
			}},
			{"additionalProperties", false}
		}}
	});
	tools.push_back({
		{"name", "search_support_library_public_code"},
		{"description", "Search keyword line-by-line inside support-library public declaration text. Prefer GetNewInf/lib2.h structured text when the library file can be resolved; otherwise it falls back to the IDE support-library info text."},
		{"inputSchema", {
			{"type", "object"},
			{"properties", {
				{"keyword", {{"type", "string"}}},
				{"index", {{"type", "integer"}, {"minimum", 0}}},
				{"name", {{"type", "string"}}},
				{"file_path", {{"type", "string"}}},
				{"limit", {{"type", "integer"}, {"minimum", 1}, {"maximum", 500}}}
			}},
			{"required", nlohmann::json::array({"keyword"})},
			{"additionalProperties", false}
		}}
	});
	tools.push_back({
		{"name", "read_support_library_public_code"},
		{"description", "Read a specific line range from one support library's cached public declaration text. Prefer md5 from search_support_library_public_code results when available; otherwise index/name/file_path can be used."},
		{"inputSchema", {
			{"type", "object"},
			{"properties", {
				{"md5", {{"type", "string"}}},
				{"index", {{"type", "integer"}, {"minimum", 0}}},
				{"name", {{"type", "string"}}},
				{"file_path", {{"type", "string"}}},
				{"start_line", {{"type", "integer"}, {"minimum", 1}}},
				{"end_line", {{"type", "integer"}, {"minimum", 1}}}
			}},
			{"required", nlohmann::json::array({"start_line", "end_line"})},
			{"additionalProperties", false}
		}}
	});
	tools.push_back({
		{"name", "get_module_public_info"},
		{"description", "Load one imported module's public interface records by module_name or module_path, primarily from offline .ec module parsing. Important: this is not normal IDE source code and only represents public-interface pseudo-reference."},
		{"inputSchema", {
			{"type", "object"},
			{"properties", {
				{"module_name", {{"type", "string"}, {"description", "Module file name or stem, for example demo.ecom or demo."}}},
				{"module_path", {{"type", "string"}}},
				{"max_records", {{"type", "integer"}, {"minimum", 1}, {"maximum", 500}}},
				{"max_strings_per_record", {{"type", "integer"}, {"minimum", 1}, {"maximum", 20}}}
			}},
			{"additionalProperties", false}
		}}
	});
	tools.push_back({
		{"name", "search_module_public_code"},
		{"description", "Search keyword line-by-line inside imported module public declaration text. Returns module, md5 and matched line numbers so the caller can read exact line ranges later. Important: this is public-interface pseudo-reference, not full module source code."},
		{"inputSchema", {
			{"type", "object"},
			{"properties", {
				{"keyword", {{"type", "string"}}},
				{"module_name", {{"type", "string"}}},
				{"module_path", {{"type", "string"}}},
				{"limit", {{"type", "integer"}, {"minimum", 1}, {"maximum", 500}}}
			}},
			{"required", nlohmann::json::array({"keyword"})},
			{"additionalProperties", false}
		}}
	});
	tools.push_back({
		{"name", "read_module_public_code"},
		{"description", "Read a specific line range from one module's cached public declaration text. Prefer md5 from search_module_public_code results when available; otherwise module_name/module_path can be used."},
		{"inputSchema", {
			{"type", "object"},
			{"properties", {
				{"md5", {{"type", "string"}}},
				{"module_name", {{"type", "string"}}},
				{"module_path", {{"type", "string"}}},
				{"start_line", {{"type", "integer"}, {"minimum", 1}}},
				{"end_line", {{"type", "integer"}, {"minimum", 1}}}
			}},
			{"required", nlohmann::json::array({"start_line", "end_line"})},
			{"additionalProperties", false}
		}}
	});
	tools.push_back({
		{"name", "list_program_items"},
		{"description", "List program tree items such as assemblies, class modules, global variables, user-defined types, DLL commands, forms and resources. Can optionally include code for each item. Important: code returned by program-tree lookup is only pseudo-code reference and may differ from the normal IDE page structure."},
		{"inputSchema", {
			{"type", "object"},
			{"properties", {
				{"kind", {{"type", "string"}, {"description", "Filter by item kind. Supported: all, assembly, class_module, global_var, user_data_type, dll_command, form, const_resource, picture_resource, sound_resource."}}},
				{"name_contains", {{"type", "string"}}},
				{"exact_name", {{"type", "string"}}},
				{"include_code", {{"type", "boolean"}}},
				{"limit", {{"type", "integer"}, {"minimum", 1}, {"maximum", 200}}}
			}},
			{"additionalProperties", false}
		}}
	});
	tools.push_back({
		{"name", "get_program_item_real_code"},
		{"description", "Get the real full source text of one program tree page by exact page_name, optionally constrained by kind. This switches the IDE current page and uses the editor's internal copy path without touching the system clipboard. Use this when you need source text matching manual Select-All plus Copy."},
		{"inputSchema", {
			{"type", "object"},
			{"properties", {
				{"page_name", {{"type", "string"}}},
				{"kind", {{"type", "string"}, {"description", "Optional kind filter: assembly, class_module, global_var, user_data_type, dll_command, form, const_resource, picture_resource, sound_resource."}}}
			}},
			{"required", nlohmann::json::array({"page_name"})},
			{"additionalProperties", false}
		}}
	});
	tools.push_back({
		{"name", "edit_program_item_code"},
		{"description", "Edit one program tree page by exact page_name. Requires a previously cached real page snapshot from get_program_item_real_code, replaces one exact old_text occurrence inside that cached page, then deletes and rewrites the full page through the editor's internal path without touching the system clipboard."},
		{"inputSchema", {
			{"type", "object"},
			{"properties", {
				{"page_name", {{"type", "string"}}},
				{"kind", {{"type", "string"}, {"description", "Optional kind filter: assembly, class_module, global_var, user_data_type, dll_command, form, const_resource, picture_resource, sound_resource."}}},
				{"old_text", {{"type", "string"}}},
				{"new_text", {{"type", "string"}}}
			}},
			{"required", nlohmann::json::array({"page_name", "old_text", "new_text"})},
			{"additionalProperties", false}
		}}
	});
	tools.push_back({
		{"name", "read_program_item_real_code"},
		{"description", "Read one real program page from cache or live editor state and optionally return only a line range view. Use refresh_cache=true when you need a fresh capture from the IDE."},
		{"inputSchema", {
			{"type", "object"},
			{"properties", {
				{"page_name", {{"type", "string"}}},
				{"kind", {{"type", "string"}, {"description", "Optional kind filter: assembly, class_module, global_var, user_data_type, dll_command, form, const_resource, picture_resource, sound_resource."}}},
				{"offset_lines", {{"type", "integer"}, {"minimum", 0}}},
				{"limit_lines", {{"type", "integer"}, {"minimum", 0}}},
				{"with_line_numbers", {{"type", "boolean"}}},
				{"refresh_cache", {{"type", "boolean"}}}
			}},
			{"required", nlohmann::json::array({"page_name"})},
			{"additionalProperties", false}
		}}
	});
	tools.push_back({
		{"name", "multi_edit_program_item_code"},
		{"description", "Apply multiple exact text edits against one cached real page, then rewrite the whole page through the internal editor path. By default this is atomic and fails if any edit does not match."},
		{"inputSchema", {
			{"type", "object"},
			{"properties", {
				{"page_name", {{"type", "string"}}},
				{"kind", {{"type", "string"}, {"description", "Optional kind filter: assembly, class_module, global_var, user_data_type, dll_command, form, const_resource, picture_resource, sound_resource."}}},
				{"edits", {{"type", "array"}, {"items", {
					{"type", "object"},
					{"properties", {
						{"old_text", {{"type", "string"}}},
						{"new_text", {{"type", "string"}}},
						{"replace_all", {{"type", "boolean"}}}
					}},
					{"required", nlohmann::json::array({"old_text", "new_text"})},
					{"additionalProperties", false}
				}}}},
				{"fail_on_unmatched", {{"type", "boolean"}}},
				{"atomic", {{"type", "boolean"}}}
			}},
			{"required", nlohmann::json::array({"page_name", "edits"})},
			{"additionalProperties", false}
		}}
	});
	tools.push_back({
		{"name", "write_program_item_real_code"},
		{"description", "Overwrite one real program page with the provided full source text using the internal full-page delete and rewrite path. Optionally provide expected_base_hash to guard against concurrent page changes."},
		{"inputSchema", {
			{"type", "object"},
			{"properties", {
				{"page_name", {{"type", "string"}}},
				{"kind", {{"type", "string"}, {"description", "Optional kind filter: assembly, class_module, global_var, user_data_type, dll_command, form, const_resource, picture_resource, sound_resource."}}},
				{"full_code", {{"type", "string"}}},
				{"expected_base_hash", {{"type", "string"}}}
			}},
			{"required", nlohmann::json::array({"page_name", "full_code"})},
			{"additionalProperties", false}
		}}
	});
	tools.push_back({
		{"name", "diff_program_item_code"},
		{"description", "Build a structured diff for one real page against a proposed full_code or text edits without writing anything back to the editor."},
		{"inputSchema", {
			{"type", "object"},
			{"properties", {
				{"page_name", {{"type", "string"}}},
				{"kind", {{"type", "string"}, {"description", "Optional kind filter: assembly, class_module, global_var, user_data_type, dll_command, form, const_resource, picture_resource, sound_resource."}}},
				{"new_code", {{"type", "string"}}},
				{"old_text", {{"type", "string"}}},
				{"new_text", {{"type", "string"}}},
				{"edits", {{"type", "array"}, {"items", {
					{"type", "object"},
					{"properties", {
						{"old_text", {{"type", "string"}}},
						{"new_text", {{"type", "string"}}},
						{"replace_all", {{"type", "boolean"}}}
					}},
					{"required", nlohmann::json::array({"old_text", "new_text"})},
					{"additionalProperties", false}
				}}}},
				{"refresh_cache", {{"type", "boolean"}}},
				{"fail_on_unmatched", {{"type", "boolean"}}}
			}},
			{"required", nlohmann::json::array({"page_name"})},
			{"additionalProperties", false}
		}}
	});
	tools.push_back({
		{"name", "restore_program_item_code_snapshot"},
		{"description", "Restore one real page from the latest cached snapshot or a specified snapshot_id. Snapshots are created automatically before successful full-page writes."},
		{"inputSchema", {
			{"type", "object"},
			{"properties", {
				{"page_name", {{"type", "string"}}},
				{"kind", {{"type", "string"}, {"description", "Optional kind filter: assembly, class_module, global_var, user_data_type, dll_command, form, const_resource, picture_resource, sound_resource."}}},
				{"snapshot_id", {{"type", "string"}}},
				{"restore_latest", {{"type", "boolean"}}}
			}},
			{"required", nlohmann::json::array({"page_name"})},
			{"additionalProperties", false}
		}}
	});
	tools.push_back({
		{"name", "search_program_item_real_code"},
		{"description", "Search inside one real program page and return exact line hits with optional context lines."},
		{"inputSchema", {
			{"type", "object"},
			{"properties", {
				{"page_name", {{"type", "string"}}},
				{"kind", {{"type", "string"}, {"description", "Optional kind filter: assembly, class_module, global_var, user_data_type, dll_command, form, const_resource, picture_resource, sound_resource."}}},
				{"keyword", {{"type", "string"}}},
				{"case_sensitive", {{"type", "boolean"}}},
				{"use_regex", {{"type", "boolean"}}},
				{"context_lines", {{"type", "integer"}, {"minimum", 0}, {"maximum", 20}}},
				{"limit", {{"type", "integer"}, {"minimum", 1}, {"maximum", 200}}},
				{"refresh_cache", {{"type", "boolean"}}}
			}},
			{"required", nlohmann::json::array({"page_name", "keyword"})},
			{"additionalProperties", false}
		}}
	});
	tools.push_back({
		{"name", "list_program_item_symbols"},
		{"description", "Parse one real page and list its symbols such as page header, subroutines, parameters, local variables and top-level variables."},
		{"inputSchema", {
			{"type", "object"},
			{"properties", {
				{"page_name", {{"type", "string"}}},
				{"kind", {{"type", "string"}, {"description", "Optional kind filter: assembly, class_module, global_var, user_data_type, dll_command, form, const_resource, picture_resource, sound_resource."}}},
				{"refresh_cache", {{"type", "boolean"}}}
			}},
			{"required", nlohmann::json::array({"page_name"})},
			{"additionalProperties", false}
		}}
	});
	tools.push_back({
		{"name", "get_symbol_real_code"},
		{"description", "Return the exact real source text for one parsed symbol inside a program page."},
		{"inputSchema", {
			{"type", "object"},
			{"properties", {
				{"page_name", {{"type", "string"}}},
				{"kind", {{"type", "string"}, {"description", "Optional kind filter: assembly, class_module, global_var, user_data_type, dll_command, form, const_resource, picture_resource, sound_resource."}}},
				{"symbol_name", {{"type", "string"}}},
				{"symbol_kind", {{"type", "string"}, {"description", "Examples: page_header, subroutine, assembly_variable, member_variable, parameter, local_variable."}}},
				{"parent_symbol_name", {{"type", "string"}}},
				{"occurrence", {{"type", "integer"}, {"minimum", 1}}},
				{"refresh_cache", {{"type", "boolean"}}}
			}},
			{"required", nlohmann::json::array({"page_name"})},
			{"additionalProperties", false}
		}}
	});
	tools.push_back({
		{"name", "edit_symbol_real_code"},
		{"description", "Replace one parsed symbol's full real source block inside a cached page, then rewrite the entire page through the internal editor path."},
		{"inputSchema", {
			{"type", "object"},
			{"properties", {
				{"page_name", {{"type", "string"}}},
				{"kind", {{"type", "string"}, {"description", "Optional kind filter: assembly, class_module, global_var, user_data_type, dll_command, form, const_resource, picture_resource, sound_resource."}}},
				{"symbol_name", {{"type", "string"}}},
				{"symbol_kind", {{"type", "string"}, {"description", "Examples: page_header, subroutine, assembly_variable, member_variable, parameter, local_variable."}}},
				{"parent_symbol_name", {{"type", "string"}}},
				{"occurrence", {{"type", "integer"}, {"minimum", 1}}},
				{"new_code", {{"type", "string"}}}
			}},
			{"required", nlohmann::json::array({"page_name", "new_code"})},
			{"additionalProperties", false}
		}}
	});
	tools.push_back({
		{"name", "insert_program_item_code_block"},
		{"description", "Insert one code block into a cached real page at the top, bottom, around a symbol, or around an exact anchor text, then rewrite the whole page through the internal editor path. If this tool returns ok=true, treat the insertion as completed and do not immediately issue extra rewrite calls on the same page unless the user explicitly asks for further cleanup or refinement."},
		{"inputSchema", {
			{"type", "object"},
			{"properties", {
				{"page_name", {{"type", "string"}}},
				{"kind", {{"type", "string"}, {"description", "Optional kind filter: assembly, class_module, global_var, user_data_type, dll_command, form, const_resource, picture_resource, sound_resource."}}},
				{"code_block", {{"type", "string"}}},
				{"mode", {{"type", "string"}, {"description", "One of: top, bottom, before_symbol, after_symbol, before_text, after_text."}}},
				{"symbol_name", {{"type", "string"}}},
				{"symbol_kind", {{"type", "string"}}},
				{"parent_symbol_name", {{"type", "string"}}},
				{"anchor_text", {{"type", "string"}}},
				{"occurrence", {{"type", "integer"}, {"minimum", 1}}}
			}},
			{"required", nlohmann::json::array({"page_name", "code_block", "mode"})},
			{"additionalProperties", false}
		}}
	});
	tools.push_back({
		{"name", "switch_to_program_item_page"},
		{"description", "Switch/open a program tree page by exact name, optionally constrained by kind. This will change the IDE current page and only activates that page; it does not fetch code."},
		{"inputSchema", {
			{"type", "object"},
			{"properties", {
				{"name", {{"type", "string"}}},
				{"kind", {{"type", "string"}, {"description", "Optional kind filter: assembly, class_module, global_var, user_data_type, dll_command, form, const_resource, picture_resource, sound_resource."}}}
			}},
			{"required", nlohmann::json::array({"name"})},
			{"additionalProperties", false}
		}}
	});
	tools.push_back({
		{"name", "search_project_keyword"},
		{"description", "Search current IDE project source hits only using the IDE hidden project search, and return matched page names, line numbers, a jump_token, and same-page occurrence metadata for each result. Use this when you only want current project hits rather than the complete unified search."},
		{"inputSchema", {
			{"type", "object"},
			{"properties", {
				{"keyword", {{"type", "string"}}},
				{"limit", {{"type", "integer"}, {"minimum", 1}, {"maximum", 200}}}
			}},
			{"required", nlohmann::json::array({"keyword"})},
			{"additionalProperties", false}
		}}
	});
	tools.push_back({
		{"name", "read_project_search_result_code"},
		{"description", "Read real page code around one IDE search hit. This jumps to the search result by jump_token, captures and caches the full current page code, then resolves the real line number by matching search_text inside the captured page. If same_text_occurrence_index is provided from search_project_keyword, it will prefer the Nth same-text hit within that page."},
		{"inputSchema", {
			{"type", "object"},
			{"properties", {
				{"jump_token", {{"type", "string"}}},
				{"search_text", {{"type", "string"}, {"description", "Pass through the text field returned by search_project_keyword for this hit."}}},
				{"same_text_occurrence_index", {{"type", "integer"}, {"minimum", 1}, {"description", "Recommended: pass through the same_text_occurrence_index field returned by search_project_keyword for this hit."}}},
				{"context_before", {{"type", "integer"}, {"minimum", 0}, {"maximum", 50}}},
				{"context_after", {{"type", "integer"}, {"minimum", 0}, {"maximum", 50}}},
				{"refresh_cache", {{"type", "boolean"}}}
			}},
			{"required", nlohmann::json::array({"jump_token", "search_text"})},
			{"additionalProperties", false}
		}}
	});
	tools.push_back({
		{"name", "jump_to_search_result"},
		{"description", "Jump to one specific search result returned by search_project_keyword using that row's jump_token. This will change the IDE current page and caret position."},
		{"inputSchema", {
			{"type", "object"},
			{"properties", {
				{"jump_token", {{"type", "string"}}}
			}},
			{"required", nlohmann::json::array({"jump_token"})},
			{"additionalProperties", false}
		}}
	});
	tools.push_back({
		{"name", "compile_with_output_path"},
		{"description", "Start compile or static compile with a specified output path and suppress the IDE system save-file dialog. Supported targets: win_exe, win_console_exe, win_dll, ecom. Note: ecom only supports non-static compile, and final compile success still needs IDE output verification."},
		{"inputSchema", {
			{"type", "object"},
			{"properties", {
				{"target", {{"type", "string"}, {"description", "One of: win_exe, win_console_exe, win_dll, ecom."}}},
				{"output_path", {{"type", "string"}}},
				{"static_compile", {{"type", "boolean"}}}
			}},
			{"required", nlohmann::json::array({"target", "output_path"})},
			{"additionalProperties", false}
		}}
	});
	tools.push_back({
		{"name", "run_powershell_command"},
		{"description", "Run one PowerShell command on the local machine after explicit user confirmation. Use it for environment inspection, file discovery or controlled automation. Avoid destructive commands unless clearly necessary."},
		{"inputSchema", {
			{"type", "object"},
			{"properties", {
				{"command", {{"type", "string"}}},
				{"working_directory", {{"type", "string"}, {"description", "Optional absolute working directory path."}}},
				{"timeout_seconds", {{"type", "integer"}, {"minimum", 1}, {"maximum", 600}}}
			}},
			{"required", nlohmann::json::array({"command"})},
			{"additionalProperties", false}
		}}
	});
	tools.push_back({
		{"name", "search_web_tavily"},
		{"description", "Search the public web via Tavily and return normalized result snippets. Prefer this when the answer depends on external up-to-date information."},
		{"inputSchema", {
			{"type", "object"},
			{"properties", {
				{"query", {{"type", "string"}}},
				{"max_results", {{"type", "integer"}, {"minimum", 1}, {"maximum", 10}}},
				{"topic", {{"type", "string"}, {"description", "Optional Tavily topic. Use general unless there is a clear reason to use another topic."}}}
			}},
			{"required", nlohmann::json::array({"query"})},
			{"additionalProperties", false}
		}}
	});
	tools.push_back({
		{"name", "fetch_url"},
		{"description", "Fetch one URL via HTTP GET and return normalized text response plus basic HTTP metadata. Use this when you need the raw response body for debugging or non-HTML text endpoints."},
		{"inputSchema", {
			{"type", "object"},
			{"properties", {
				{"url", {{"type", "string"}}},
				{"timeout_seconds", {{"type", "integer"}, {"minimum", 1}, {"maximum", 300}}},
				{"max_bytes", {{"type", "integer"}, {"minimum", 4096}, {"maximum", 2097152}}}
			}},
			{"required", nlohmann::json::array({"url"})},
			{"additionalProperties", false}
		}}
	});
	tools.push_back({
		{"name", "extract_web_document"},
		{"description", "Fetch one web page or text document and extract readable plain-text content, title and a small set of links. Prefer this over raw fetch_url when reading docs or articles."},
		{"inputSchema", {
			{"type", "object"},
			{"properties", {
				{"url", {{"type", "string"}}},
				{"timeout_seconds", {{"type", "integer"}, {"minimum", 1}, {"maximum", 300}}},
				{"max_bytes", {{"type", "integer"}, {"minimum", 4096}, {"maximum", 2097152}}}
			}},
			{"required", nlohmann::json::array({"url"})},
			{"additionalProperties", false}
		}}
	});
	return tools;
}

nlohmann::json BuildChatToolDefinitions()
{
	const nlohmann::json catalog = BuildPublicToolCatalog();
	nlohmann::json tools = nlohmann::json::array();
	for (const auto& item : catalog) {
		tools.push_back({
			{"type", "function"},
			{"function", {
				{"name", item.value("name", "")},
				{"description", item.value("description", "")},
				{"parameters", item.contains("inputSchema") ? item["inputSchema"] : nlohmann::json::object()}
			}}
		});
	}
	return tools;
}

std::string BuildChatSystemPrompt(const AISettings& settings)
{
	std::string projectName;
	if (!AIService::Trim(g_nowOpenSourceFilePath).empty()) {
		try {
			const std::filesystem::path sourcePath(g_nowOpenSourceFilePath);
			projectName = sourcePath.stem().string();
			if (projectName.empty()) {
				projectName = sourcePath.filename().string();
			}
		}
		catch (...) {
			projectName.clear();
		}
	}

	const std::string projectType = DetectProjectTypeText();
	std::string prompt =
		"你是 AutoLinker 内置的易语言项目助手。\n"
		"优先使用工具获取准确上下文，不要臆测当前页面、模块、支持库、搜索结果或代码内容。\n\n"
		"当前项目名称：" + (projectName.empty() ? std::string("未知") : projectName) + "\n\n"
		"当前项目类型：" + projectType + "\n\n"
		"易语言基础约定：\n"
		"- 以 # 开头的标识通常表示常量。\n"
		"- 以 . 开头的是易语言系统指令/关键字，例如 .版本、.程序集、.子程序、.参数、.常量、.DLL声明、.数据、.局部变量、.如果、.如果真、.否则。\n"
		"- 单引号 ' 开头表示注释。\n"
		"- 真 / 假 是布尔值。\n"
		"- 赋值常写作 `变量 ＝ 值`，不要误写成 C/C++ 风格的 `=`。\n"
		"- 自增通常写作 `a ＝ a ＋ 1`，不要写 `a++`。\n"
		"- 返回常见写法是 `返回 (...)`。\n"
		"- 全角中文标点和全角运算符在代码里较常见，分析时不要误判。\n\n"
		"常见流程控制示例：\n"
		"1) if：\n"
		".如果真 ()\n"
		"    a ＝ 0\n"
		".如果真结束\n\n"
		"2) if else：\n"
		".如果 (a ＝ 0)\n"
		"    a ＝ 1\n"
		".否则\n"
		"    b ＝ 1\n"
		".如果结束\n\n"
		"3) 计次循环：\n"
		".计次循环首 (count, i)\n"
		"    a ＝ a ＋ 1\n"
		".计次循环尾 ()\n\n"
		"4) 变量循环：\n"
		".变量循环首 (1, 100, 1, i)\n"
		"    输出调试文本 (i)\n"
		".变量循环尾 ()\n\n"
		"5) 循环判断：\n"
		".循环判断首 ()\n"
		"    输出调试文本 (“这么写会是死循环”)\n"
		".循环判断尾 (真)\n\n"
		"6) 多分支判断：\n"
		".判断开始 (b ＝ 1)\n"
		"    a ＝ 1\n"
		".判断 (b ＝ 2)\n"
		"    a ＝ 1\n"
		".判断 (b ＝ 3)\n"
		"    a ＝ 1\n"
		".判断 (b ＝ 4)\n"
		"    a ＝ 1\n"
		".默认\n"
		"    def ＝ 1\n"
		".判断结束\n\n"
		"工具使用规则：\n"
		"1) 需要当前页完整代码时调用 get_current_page_code；如果还要页名与类型，调用 get_current_page_info。\n"
		"2) 只想知道当前打开页是谁，不需要整页代码时，优先调用 get_current_page_info；需要当前源码路径、MCP 端口、进程路径等实例级信息时，调用 get_current_eide_info。\n"
		"3) 本机有多个易语言实例时，先用 list_local_mcp_instances，再通过 call_local_mcp_instance_tool 转发到目标实例；不要臆测端口。\n"
		"4) 支持库相关信息：先用 list_support_libraries，再按需用 get_support_library_info；若需要完全搜索当前工程源码命中、模块公开声明、支持库公开声明，优先用 search_public_code；若只查支持库也可用 search_support_library_public_code / read_support_library_public_code。\n"
		"5) 模块公开信息：先用 list_imported_modules，再按需用 get_module_public_info；若需要完全搜索当前工程源码命中、模块公开声明、支持库公开声明，优先用 search_public_code；若只查模块也可用 search_module_public_code / read_module_public_code。\n"
		"6) 程序树页面：先用 list_program_items 定位页面，再按需用 get_program_item_real_code 或 switch_to_program_item_page。\n"
		"6.1) list_program_items 附带的代码仍只是伪代码参考；需要某个页面的真实整页源码时，用 get_program_item_real_code。\n"
		"6.2) 需要分页查看或从缓存读取真实源码时，用 read_program_item_real_code。\n"
		"6.3) 需要真正改写某个页面源码时，先调用 get_program_item_real_code 或 read_program_item_real_code 建立缓存，再按需用 edit_program_item_code / multi_edit_program_item_code / write_program_item_real_code。\n"
		"6.4) 需要预览改动而不写回时，用 diff_program_item_code。\n"
		"6.5) 需要按符号操作真实源码时，用 list_program_item_symbols / get_symbol_real_code / edit_symbol_real_code / insert_program_item_code_block。\n"
		"6.6) 需要在真实页内做精确搜索或回滚最近写入时，用 search_program_item_real_code / restore_program_item_code_snapshot。\n"
		"7) 需要完全搜索当前工程源码命中、模块公开声明、支持库公开声明时，优先用 search_public_code；只有在只想搜索当前 IDE 工程源码命中时，才用 search_project_keyword。若要基于某个工程命中结果读取真实页代码行范围，优先用 read_project_search_result_code；仅在需要单纯跳转时再用 jump_to_search_result。\n"
		"8) read_project_search_result_code、jump_to_search_result、switch_to_program_item_page、get_program_item_real_code 都会改变 IDE 当前页面，调用前要意识到页面会被切走。\n"
		"9) 通过搜索、程序树、模块公开信息、支持库公开信息拿到的代码或文本，多数只是伪代码 / 公共接口参考，不一定等于 IDE 正常编辑页。\n"
		"10) 需要无弹窗编译时调用 compile_with_output_path。它会指定输出路径并拦截系统保存对话框，支持模块工程编译为 ec，以及窗口程序 / 控制台程序 / DLL 的编译与静态编译；最终是否编译成功仍要结合 IDE 输出或产物确认。\n"
		"11) 需要自动整页回写真实源码时优先使用真实页工具，不要退回伪代码工具。\n"
		"12) 需要联网查实时信息、文档、网页摘要时调用 search_web_tavily。\n"
		"13) 已经拿到具体文档 URL 时，优先调用 extract_web_document 读取正文；只有在需要看原始响应时再调用 fetch_url。\n"
		"14) 需要在本机查环境、查文件、执行受控自动化时调用 run_powershell_command；它每次都会向用户确认，命令要尽量小、明确、可解释。\n"
		"15) run_powershell_command 被用户取消后，不要机械重试，应改为解释下一步或换别的工具。\n"
		"16) 工具失败时先分析失败原因并换更合适的工具，不要机械重试同一个调用；真实页写工具一旦已经返回 ok=true，就默认停止，不要立刻对同一页继续追加无必要的二次写回。\n"
		"17) 不要要求用户手动补上下文，优先自己通过工具获取。\n";
	const std::string extraPrompt = AIService::Trim(settings.extraSystemPrompt);
	if (!extraPrompt.empty()) {
		prompt += "\n附加系统提示：\n";
		prompt += extraPrompt;
	}
	return prompt;
}

std::string UrlEncode(const std::string& value)
{
	static constexpr char kHex[] = "0123456789ABCDEF";
	std::string encoded;
	encoded.reserve(value.size() + 16);
	for (unsigned char c : value) {
		if ((c >= 'a' && c <= 'z') ||
			(c >= 'A' && c <= 'Z') ||
			(c >= '0' && c <= '9') ||
			c == '-' || c == '_' || c == '.' || c == '~') {
			encoded.push_back(static_cast<char>(c));
			continue;
		}
		encoded.push_back('%');
		encoded.push_back(kHex[(c >> 4) & 0x0F]);
		encoded.push_back(kHex[c & 0x0F]);
	}
	return encoded;
}

std::string AppendQueryParam(std::string url, const std::string& key, const std::string& value)
{
	if (key.empty()) {
		return url;
	}
	const char sep = (url.find('?') == std::string::npos) ? '?' : '&';
	url.push_back(sep);
	url += UrlEncode(key);
	url += "=";
	url += UrlEncode(value);
	return url;
}

std::string ReplaceSuffixIfPresent(const std::string& text, const std::string& oldSuffix, const std::string& newSuffix)
{
	if (!EndsWithInsensitive(text, oldSuffix)) {
		return text;
	}
	return text.substr(0, text.size() - oldSuffix.size()) + newSuffix;
}

std::string BuildClaudeEndpoint(const std::string& baseUrl)
{
	std::string url = AIService::Trim(baseUrl);
	while (!url.empty() && url.back() == '/') {
		url.pop_back();
	}
	if (EndsWithInsensitive(url, "/v1/messages")) {
		return url;
	}
	if (EndsWithInsensitive(url, "/v1")) {
		return url + "/messages";
	}
	return url + "/v1/messages";
}

std::string BuildGeminiEndpoint(const std::string& baseUrl, const std::string& model, bool stream)
{
	std::string url = AIService::Trim(baseUrl);
	while (!url.empty() && url.back() == '/') {
		url.pop_back();
	}

	const std::string suffix = stream ? ":streamGenerateContent" : ":generateContent";
	const std::string otherSuffix = stream ? ":generateContent" : ":streamGenerateContent";

	url = ReplaceSuffixIfPresent(url, otherSuffix, suffix);
	if (EndsWithInsensitive(url, suffix)) {
		return stream ? AppendQueryParam(url, "alt", "sse") : url;
	}

	if (url.find("/models/") != std::string::npos) {
		url += suffix;
		return stream ? AppendQueryParam(url, "alt", "sse") : url;
	}

	if (EndsWithInsensitive(url, "/v1beta") || EndsWithInsensitive(url, "/v1")) {
		url += "/models/" + UrlEncode(model) + suffix;
		return stream ? AppendQueryParam(url, "alt", "sse") : url;
	}

	url += "/v1beta/models/" + UrlEncode(model) + suffix;
	return stream ? AppendQueryParam(url, "alt", "sse") : url;
}

std::string BuildOpenAIHeaders(const AISettings& settings)
{
	return
		"Content-Type: application/json\r\n"
		"Authorization: Bearer " + settings.apiKey + "\r\n";
}

std::string BuildClaudeHeaders(const AISettings& settings)
{
	return
		"Content-Type: application/json\r\n"
		"x-api-key: " + settings.apiKey + "\r\n"
		"anthropic-version: 2023-06-01\r\n";
}

std::string BuildJsonHeadersOnly()
{
	return "Content-Type: application/json\r\n";
}

nlohmann::json BuildClaudeTools()
{
	nlohmann::json out = nlohmann::json::array();
	const nlohmann::json openAiTools = BuildChatToolDefinitions();
	for (const auto& tool : openAiTools) {
		if (!tool.contains("function") || !tool["function"].is_object()) {
			continue;
		}
		const nlohmann::json& fn = tool["function"];
		out.push_back({
			{"name", fn.value("name", "")},
			{"description", fn.value("description", "")},
			{"input_schema", fn.value("parameters", nlohmann::json::object())}
		});
	}
	return out;
}

nlohmann::json BuildGeminiTools()
{
	nlohmann::json declarations = nlohmann::json::array();
	const nlohmann::json openAiTools = BuildChatToolDefinitions();
	for (const auto& tool : openAiTools) {
		if (!tool.contains("function") || !tool["function"].is_object()) {
			continue;
		}
		const nlohmann::json& fn = tool["function"];
		declarations.push_back({
			{"name", fn.value("name", "")},
			{"description", fn.value("description", "")},
			{"parameters", fn.value("parameters", nlohmann::json::object())}
		});
	}
	return nlohmann::json::array({ {{"functionDeclarations", declarations}} });
}

std::string ParseErrorMessageUtf8(const nlohmann::json& parsed)
{
	if (!parsed.contains("error")) {
		return std::string();
	}
	const auto& errorNode = parsed["error"];
	if (errorNode.is_object() && errorNode.contains("message") && errorNode["message"].is_string()) {
		return errorNode["message"].get<std::string>();
	}
	if (errorNode.is_string()) {
		return errorNode.get<std::string>();
	}
	return std::string();
}

std::string ExtractClaudeTextUtf8(const nlohmann::json& parsed)
{
	if (!parsed.contains("content") || !parsed["content"].is_array()) {
		return std::string();
	}
	std::string textUtf8;
	for (const auto& item : parsed["content"]) {
		if (!item.is_object()) {
			continue;
		}
		if (item.value("type", std::string()) == "text" && item.contains("text") && item["text"].is_string()) {
			textUtf8 += item["text"].get<std::string>();
		}
	}
	return textUtf8;
}

std::string ExtractGeminiTextUtf8(const nlohmann::json& parsed)
{
	if (!parsed.contains("candidates") || !parsed["candidates"].is_array() || parsed["candidates"].empty()) {
		return std::string();
	}
	const auto& candidate = parsed["candidates"][0];
	if (!candidate.contains("content") || !candidate["content"].is_object()) {
		return std::string();
	}
	const auto& content = candidate["content"];
	if (!content.contains("parts") || !content["parts"].is_array()) {
		return std::string();
	}
	std::string textUtf8;
	for (const auto& part : content["parts"]) {
		if (part.is_object() && part.contains("text") && part["text"].is_string()) {
			textUtf8 += part["text"].get<std::string>();
		}
	}
	return textUtf8;
}

struct ClaudeToolCall {
	std::string id;
	std::string name;
	std::string argumentsUtf8;
};

struct GeminiToolCall {
	std::string name;
	std::string argumentsUtf8;
};

std::vector<ClaudeToolCall> ExtractClaudeToolCalls(const nlohmann::json& parsed)
{
	std::vector<ClaudeToolCall> calls;
	if (!parsed.contains("content") || !parsed["content"].is_array()) {
		return calls;
	}

	for (const auto& item : parsed["content"]) {
		if (!item.is_object() || item.value("type", std::string()) != "tool_use") {
			continue;
		}
		ClaudeToolCall call;
		call.id = item.value("id", "");
		call.name = item.value("name", "");
		if (item.contains("input")) {
			call.argumentsUtf8 = item["input"].dump();
		}
		else {
			call.argumentsUtf8 = "{}";
		}
		calls.push_back(std::move(call));
	}
	return calls;
}

std::vector<GeminiToolCall> ExtractGeminiToolCalls(const nlohmann::json& parsed)
{
	std::vector<GeminiToolCall> calls;
	if (!parsed.contains("candidates") || !parsed["candidates"].is_array() || parsed["candidates"].empty()) {
		return calls;
	}
	const auto& candidate = parsed["candidates"][0];
	if (!candidate.contains("content") || !candidate["content"].is_object()) {
		return calls;
	}
	const auto& content = candidate["content"];
	if (!content.contains("parts") || !content["parts"].is_array()) {
		return calls;
	}

	for (const auto& part : content["parts"]) {
		if (!part.is_object() || !part.contains("functionCall") || !part["functionCall"].is_object()) {
			continue;
		}
		const auto& fn = part["functionCall"];
		GeminiToolCall call;
		call.name = fn.value("name", "");
		if (fn.contains("args")) {
			call.argumentsUtf8 = fn["args"].dump();
		}
		else {
			call.argumentsUtf8 = "{}";
		}
		calls.push_back(std::move(call));
	}
	return calls;
}

AIResult ExecuteTaskClaude(const std::string& systemPrompt, const std::string& inputText, const AISettings& settings)
{
	AIResult result = {};
	const std::string endpoint = BuildClaudeEndpoint(settings.baseUrl);

	nlohmann::json requestBody;
	requestBody["model"] = LocalToUtf8(settings.model);
	requestBody["max_tokens"] = 4096;
	requestBody["temperature"] = settings.temperature;
	requestBody["system"] = LocalToUtf8(systemPrompt);
	requestBody["messages"] = nlohmann::json::array({
		{
			{"role", "user"},
			{"content", nlohmann::json::array({ {{"type", "text"}, {"text", LocalToUtf8(inputText)}} })}
		}
	});

	const auto [responseBody, statusCode] = PerformPostRequestWithRetry(
		endpoint,
		requestBody.dump(),
		BuildClaudeHeaders(settings),
		settings.timeoutMs,
		false,
		false,
		"claude-task");
	result.httpStatus = statusCode;
	if (statusCode < 200 || statusCode >= 300) {
		result.error = std::format("HTTP {}: {}", statusCode, TruncateForLog(responseBody));
		return result;
	}

	try {
		const nlohmann::json parsed = nlohmann::json::parse(responseBody);
		const std::string errUtf8 = ParseErrorMessageUtf8(parsed);
		if (!errUtf8.empty()) {
			result.error = Utf8ToLocal(errUtf8);
			return result;
		}
		const std::string textUtf8 = ExtractClaudeTextUtf8(parsed);
		if (textUtf8.empty()) {
			result.error = "Claude response content is empty";
			return result;
		}
		result.ok = true;
		result.content = Utf8ToLocal(textUtf8);
		return result;
	}
	catch (const std::exception& ex) {
		result.error = std::string("Failed to parse Claude response: ") + ex.what();
		return result;
	}
}

AIResult ExecuteTaskGemini(const std::string& systemPrompt, const std::string& inputText, const AISettings& settings)
{
	AIResult result = {};
	std::string endpoint = BuildGeminiEndpoint(settings.baseUrl, LocalToUtf8(settings.model), false);
	endpoint = AppendQueryParam(endpoint, "key", settings.apiKey);

	nlohmann::json requestBody;
	requestBody["system_instruction"] = {
		{"parts", nlohmann::json::array({ {{"text", LocalToUtf8(systemPrompt)}} })}
	};
	requestBody["generationConfig"] = { {"temperature", settings.temperature} };
	requestBody["contents"] = nlohmann::json::array({
		{
			{"role", "user"},
			{"parts", nlohmann::json::array({ {{"text", LocalToUtf8(inputText)}} })}
		}
	});

	const auto [responseBody, statusCode] = PerformPostRequestWithRetry(
		endpoint,
		requestBody.dump(),
		BuildJsonHeadersOnly(),
		settings.timeoutMs,
		false,
		false,
		"gemini-task");
	result.httpStatus = statusCode;
	if (statusCode < 200 || statusCode >= 300) {
		result.error = std::format("HTTP {}: {}", statusCode, TruncateForLog(responseBody));
		return result;
	}

	try {
		const nlohmann::json parsed = nlohmann::json::parse(responseBody);
		const std::string errUtf8 = ParseErrorMessageUtf8(parsed);
		if (!errUtf8.empty()) {
			result.error = Utf8ToLocal(errUtf8);
			return result;
		}
		const std::string textUtf8 = ExtractGeminiTextUtf8(parsed);
		if (textUtf8.empty()) {
			result.error = "Gemini response content is empty";
			return result;
		}
		result.ok = true;
		result.content = Utf8ToLocal(textUtf8);
		return result;
	}
	catch (const std::exception& ex) {
		result.error = std::string("Failed to parse Gemini response: ") + ex.what();
		return result;
	}
}

AIChatResult ExecuteChatWithToolsClaude(
	const std::vector<AIChatMessage>& contextMessages,
	const AISettings& settings,
	const std::function<std::string(const std::string& toolName, const std::string& argumentsJson, bool& outOk)>& toolCallback,
	const std::function<void(const std::string& deltaText)>& streamCallback)
{
	AIChatResult result = {};
	const std::string endpoint = BuildClaudeEndpoint(settings.baseUrl);
	const nlohmann::json tools = BuildClaudeTools();

	std::string systemUtf8 = LocalToUtf8(BuildChatSystemPrompt(settings));
	nlohmann::json messages = nlohmann::json::array();
	for (const AIChatMessage& msg : contextMessages) {
		const std::string role = ToLowerAsciiCopy(AIService::Trim(msg.role));
		if (role == "system") {
			systemUtf8 += "\n\n";
			systemUtf8 += LocalToUtf8(msg.content);
			continue;
		}
		if (role != "user" && role != "assistant") {
			continue;
		}
		messages.push_back({
			{"role", role},
			{"content", nlohmann::json::array({
				{{"type", "text"}, {"text", LocalToUtf8(msg.content)}}
			})}
		});
	}

	const int maxToolRounds = GetMaxToolRounds(settings);
	for (int round = 0; round < maxToolRounds; ++round) {
		nlohmann::json requestBody;
		requestBody["model"] = LocalToUtf8(settings.model);
		requestBody["max_tokens"] = 4096;
		requestBody["temperature"] = settings.temperature;
		requestBody["system"] = systemUtf8;
		requestBody["messages"] = messages;
		requestBody["tools"] = tools;
		requestBody["tool_choice"] = { {"type", "auto"} };
		requestBody["stream"] = false;

		const auto [responseBody, statusCode] = PerformPostRequestWithRetry(
			endpoint,
			requestBody.dump(),
			BuildClaudeHeaders(settings),
			settings.timeoutMs,
			false,
			false,
			"claude-chat");
		result.httpStatus = statusCode;
		if (statusCode < 200 || statusCode >= 300) {
			result.error = std::format("HTTP {}: {}", statusCode, TruncateForLog(responseBody));
			return result;
		}

		nlohmann::json parsed;
		try {
			parsed = nlohmann::json::parse(responseBody);
		}
		catch (const std::exception& ex) {
			result.error = std::string("Failed to parse Claude response: ") + ex.what();
			return result;
		}

		const std::string errUtf8 = ParseErrorMessageUtf8(parsed);
		if (!errUtf8.empty()) {
			result.error = Utf8ToLocal(errUtf8);
			return result;
		}

		const std::vector<ClaudeToolCall> toolCalls = ExtractClaudeToolCalls(parsed);
		const std::string textUtf8 = ExtractClaudeTextUtf8(parsed);
		if (toolCalls.empty()) {
			if (textUtf8.empty()) {
				result.error = "Claude response content is empty";
				return result;
			}
			result.ok = true;
			result.content = Utf8ToLocal(textUtf8);
			if (streamCallback) {
				streamCallback(result.content);
			}
			return result;
		}

		if (parsed.contains("content") && parsed["content"].is_array()) {
			messages.push_back({
				{"role", "assistant"},
				{"content", parsed["content"]}
			});
		}

		for (size_t i = 0; i < toolCalls.size(); ++i) {
			const ClaudeToolCall& call = toolCalls[i];
			const std::string callId = call.id.empty()
				? std::format("toolu_auto_{}_{}", round + 1, i + 1)
				: call.id;

			bool toolOk = false;
			std::string toolResultLocal;
			if (toolCallback) {
				toolResultLocal = toolCallback(call.name, call.argumentsUtf8, toolOk);
			}
			else {
				toolResultLocal = R"({"ok":false,"error":"tool callback not set"})";
			}
			const CompactToolResultPayload compactPayload = BuildCompactToolResultPayload(call.name, toolResultLocal);

			AIChatToolEvent evt = {};
			evt.name = call.name;
			evt.argumentsJson = Utf8ToLocal(call.argumentsUtf8);
			evt.resultJson = toolResultLocal;
			evt.ok = toolOk;
			result.toolEvents.push_back(std::move(evt));

			messages.push_back({
				{"role", "user"},
				{"content", nlohmann::json::array({
					{
						{"type", "tool_result"},
						{"tool_use_id", callId},
						{"content", compactPayload.textUtf8}
					}
				})}
			});
		}
	}

	result.error = BuildToolRoundsExceededError(maxToolRounds, result.toolEvents);
	return result;
}

AIChatResult ExecuteChatWithToolsGemini(
	const std::vector<AIChatMessage>& contextMessages,
	const AISettings& settings,
	const std::function<std::string(const std::string& toolName, const std::string& argumentsJson, bool& outOk)>& toolCallback,
	const std::function<void(const std::string& deltaText)>& streamCallback)
{
	AIChatResult result = {};
	std::string endpoint = BuildGeminiEndpoint(settings.baseUrl, LocalToUtf8(settings.model), false);
	endpoint = AppendQueryParam(endpoint, "key", settings.apiKey);
	const nlohmann::json tools = BuildGeminiTools();

	std::string systemUtf8 = LocalToUtf8(BuildChatSystemPrompt(settings));
	nlohmann::json contents = nlohmann::json::array();
	for (const AIChatMessage& msg : contextMessages) {
		const std::string role = ToLowerAsciiCopy(AIService::Trim(msg.role));
		if (role == "system") {
			systemUtf8 += "\n\n";
			systemUtf8 += LocalToUtf8(msg.content);
			continue;
		}
		if (role != "user" && role != "assistant") {
			continue;
		}
		contents.push_back({
			{"role", role == "assistant" ? "model" : "user"},
			{"parts", nlohmann::json::array({
				{{"text", LocalToUtf8(msg.content)}}
			})}
		});
	}

	const int maxToolRounds = GetMaxToolRounds(settings);
	for (int round = 0; round < maxToolRounds; ++round) {
		nlohmann::json requestBody;
		requestBody["system_instruction"] = {
			{"parts", nlohmann::json::array({ {{"text", systemUtf8}} })}
		};
		requestBody["generationConfig"] = { {"temperature", settings.temperature} };
		requestBody["contents"] = contents;
		requestBody["tools"] = tools;

		const auto [responseBody, statusCode] = PerformPostRequestWithRetry(
			endpoint,
			requestBody.dump(),
			BuildJsonHeadersOnly(),
			settings.timeoutMs,
			false,
			false,
			"gemini-chat");
		result.httpStatus = statusCode;
		if (statusCode < 200 || statusCode >= 300) {
			result.error = std::format("HTTP {}: {}", statusCode, TruncateForLog(responseBody));
			return result;
		}

		nlohmann::json parsed;
		try {
			parsed = nlohmann::json::parse(responseBody);
		}
		catch (const std::exception& ex) {
			result.error = std::string("Failed to parse Gemini response: ") + ex.what();
			return result;
		}

		const std::string errUtf8 = ParseErrorMessageUtf8(parsed);
		if (!errUtf8.empty()) {
			result.error = Utf8ToLocal(errUtf8);
			return result;
		}

		if (!parsed.contains("candidates") || !parsed["candidates"].is_array() || parsed["candidates"].empty()) {
			result.error = "Gemini response candidates is empty";
			return result;
		}

		const auto& candidate = parsed["candidates"][0];
		if (!candidate.contains("content") || !candidate["content"].is_object()) {
			result.error = "Gemini response content missing";
			return result;
		}
		const auto& candidateContent = candidate["content"];

		const std::vector<GeminiToolCall> toolCalls = ExtractGeminiToolCalls(parsed);
		const std::string textUtf8 = ExtractGeminiTextUtf8(parsed);
		if (toolCalls.empty()) {
			if (textUtf8.empty()) {
				result.error = "Gemini response content is empty";
				return result;
			}
			result.ok = true;
			result.content = Utf8ToLocal(textUtf8);
			if (streamCallback) {
				streamCallback(result.content);
			}
			return result;
		}

		contents.push_back(candidateContent);

		for (const GeminiToolCall& call : toolCalls) {
			bool toolOk = false;
			std::string toolResultLocal;
			if (toolCallback) {
				toolResultLocal = toolCallback(call.name, call.argumentsUtf8, toolOk);
			}
			else {
				toolResultLocal = R"({"ok":false,"error":"tool callback not set"})";
			}
			const CompactToolResultPayload compactPayload = BuildCompactToolResultPayload(call.name, toolResultLocal);

			AIChatToolEvent evt = {};
			evt.name = call.name;
			evt.argumentsJson = Utf8ToLocal(call.argumentsUtf8);
			evt.resultJson = toolResultLocal;
			evt.ok = toolOk;
			result.toolEvents.push_back(std::move(evt));

			contents.push_back({
				{"role", "user"},
				{"parts", nlohmann::json::array({
					{
						{"functionResponse", {
							{"name", call.name},
							{"response", compactPayload.jsonValue}
						}}
					}
				})}
			});
		}
	}

	result.error = BuildToolRoundsExceededError(maxToolRounds, result.toolEvents);
	return result;
}
} // namespace

bool AIService::LoadSettings(ConfigManager& config, AISettings& outSettings)
{
	outSettings = {};
	outSettings.protocolType = ParseProtocolType(config.getValue("ai.protocol_type"));
	outSettings.baseUrl = config.getValue("ai.base_url");
	outSettings.apiKey = config.getValue("ai.api_key");
	outSettings.model = config.getValue("ai.model");
	outSettings.extraSystemPrompt = config.getValue("ai.system_prompt_extra");
	outSettings.tavilyApiKey = config.getValue("ai.tavily_api_key");

	const std::string timeoutValue = config.getValue("ai.timeout_ms");
	if (!timeoutValue.empty()) {
		try {
			outSettings.timeoutMs = (std::max)(1000, std::stoi(timeoutValue));
		}
		catch (...) {
			outSettings.timeoutMs = 120000;
		}
	}

	const std::string temperatureValue = config.getValue("ai.temperature");
	if (!temperatureValue.empty()) {
		try {
			outSettings.temperature = std::stod(temperatureValue);
		}
		catch (...) {
			outSettings.temperature = 0.2;
		}
	}

	const std::string maxToolRoundsValue = config.getValue("ai.max_tool_rounds");
	if (!maxToolRoundsValue.empty()) {
		try {
			outSettings.maxToolRounds = (std::clamp)(std::stoi(maxToolRoundsValue), 4, 64);
		}
		catch (...) {
			outSettings.maxToolRounds = 48;
		}
	}

	return true;
}

void AIService::SaveSettings(ConfigManager& config, const AISettings& settings)
{
	config.setValue("ai.protocol_type", ProtocolTypeToString(settings.protocolType));
	config.setValue("ai.base_url", settings.baseUrl);
	config.setValue("ai.api_key", settings.apiKey);
	config.setValue("ai.model", settings.model);
	config.setValue("ai.system_prompt_extra", settings.extraSystemPrompt);
	config.setValue("ai.tavily_api_key", settings.tavilyApiKey);
	config.setValue("ai.timeout_ms", std::to_string(settings.timeoutMs));
	config.setValue("ai.max_tool_rounds", std::to_string((std::clamp)(settings.maxToolRounds, 4, 64)));
	config.setValue("ai.temperature", std::format("{:.2f}", settings.temperature));
}

bool AIService::HasRequiredSettings(const AISettings& settings, std::string& outMissingField)
{
	if (Trim(settings.baseUrl).empty()) {
		outMissingField = "baseUrl";
		return false;
	}
	if (Trim(settings.apiKey).empty()) {
		outMissingField = "apiKey";
		return false;
	}
	if (Trim(settings.model).empty()) {
		outMissingField = "model";
		return false;
	}
	outMissingField.clear();
	return true;
}

AIProtocolType AIService::ParseProtocolType(const std::string& text)
{
	const std::string v = ToLowerAsciiCopy(Trim(text));
	if (v == "gemini") {
		return AIProtocolType::Gemini;
	}
	if (v == "claude") {
		return AIProtocolType::Claude;
	}
	return AIProtocolType::OpenAI;
}

std::string AIService::ProtocolTypeToString(AIProtocolType protocolType)
{
	switch (protocolType) {
	case AIProtocolType::Gemini:
		return "gemini";
	case AIProtocolType::Claude:
		return "claude";
	case AIProtocolType::OpenAI:
	default:
		return "openai";
	}
}

std::string AIService::ProtocolTypeDisplayName(AIProtocolType protocolType)
{
	switch (protocolType) {
	case AIProtocolType::Gemini:
		return "Gemini";
	case AIProtocolType::Claude:
		return "Claude";
	case AIProtocolType::OpenAI:
	default:
		return "OpenAI";
	}
}

std::string AIService::BuildTaskDisplayName(AITaskKind kind)
{
	switch (kind)
	{
	case AITaskKind::OptimizeFunction:
		return "AI优化函数";
	case AITaskKind::AddCommentsToFunction:
		return "AI为当前函数添加注释";
	case AITaskKind::TranslateFunctionAndVariables:
		return "AI翻译当前函数+变量名";
	case AITaskKind::TranslateText:
		return "AI翻译选中文本";
	case AITaskKind::AddByCurrentPageType:
		return "AI按当前页类型添加代码";
	default:
		return "AI任务";
	}
}

AIResult AIService::ExecuteTask(AITaskKind kind, const std::string& inputText, const AISettings& settings)
{
	AIResult result = {};
	std::string missingField;
	if (!HasRequiredSettings(settings, missingField)) {
		result.error = "AI settings missing: " + missingField;
		return result;
	}

	const std::string systemPrompt = BuildSystemPrompt(kind, settings);
	if (settings.protocolType == AIProtocolType::Claude) {
		return ExecuteTaskClaude(systemPrompt, inputText, settings);
	}
	if (settings.protocolType == AIProtocolType::Gemini) {
		return ExecuteTaskGemini(systemPrompt, inputText, settings);
	}

	const std::string modelUtf8 = LocalToUtf8(settings.model);
	const std::string systemPromptUtf8 = LocalToUtf8(systemPrompt);
	const std::string inputTextUtf8 = LocalToUtf8(inputText);

	nlohmann::json requestBody;
	requestBody["model"] = modelUtf8;
	requestBody["temperature"] = settings.temperature;
	requestBody["stream"] = false;
	requestBody["messages"] = nlohmann::json::array({
		{
			{"role", "system"},
			{"content", systemPromptUtf8}
		},
		{
			{"role", "user"},
			{"content", inputTextUtf8}
		}
	});

	const std::string endpoint = BuildEndpoint(settings.baseUrl);
	const std::string headers = BuildOpenAIHeaders(settings);

	std::string requestBodyText;
	try {
		requestBodyText = requestBody.dump();
	}
	catch (const std::exception& ex) {
		result.error = std::string("Failed to build AI request JSON: ") + ex.what();
		return result;
	}

	const auto [responseBody, statusCode] =
		PerformPostRequestWithRetry(endpoint, requestBodyText, headers, settings.timeoutMs, false, false, "openai-task");
	result.httpStatus = statusCode;

	if (statusCode < 200 || statusCode >= 300) {
		result.error = std::format("HTTP {}: {}", statusCode, TruncateForLog(responseBody));
		return result;
	}

	try {
		const nlohmann::json parsed = nlohmann::json::parse(responseBody);
		if (parsed.contains("choices") && parsed["choices"].is_array() && !parsed["choices"].empty()) {
			const nlohmann::json& choice = parsed["choices"][0];
			if (choice.contains("message") && choice["message"].contains("content")) {
				if (choice["message"]["content"].is_string()) {
					result.ok = true;
					result.content = Utf8ToLocal(choice["message"]["content"].get<std::string>());
					return result;
				}
				if (choice["message"]["content"].is_array()) {
					std::string merged;
					for (const auto& item : choice["message"]["content"]) {
						if (item.is_string()) {
							merged += item.get<std::string>();
							continue;
						}
						if (item.is_object() && item.contains("text") && item["text"].is_string()) {
							merged += item["text"].get<std::string>();
						}
					}
					if (!merged.empty()) {
						result.ok = true;
						result.content = Utf8ToLocal(merged);
						return result;
					}
				}
			}
		}

		if (parsed.contains("error") && parsed["error"].contains("message") && parsed["error"]["message"].is_string()) {
			result.error = Utf8ToLocal(parsed["error"]["message"].get<std::string>());
			return result;
		}

		result.error = "AI response does not match expected chat/completions schema";
	}
	catch (const std::exception& ex) {
		result.error = std::string("Failed to parse AI response: ") + ex.what();
	}

	return result;
}

AIChatResult AIService::ExecuteChatWithTools(
	const std::vector<AIChatMessage>& contextMessages,
	const AISettings& settings,
	const std::function<std::string(const std::string& toolName, const std::string& argumentsJson, bool& outOk)>& toolCallback,
	const std::function<void(const std::string& deltaText)>& streamCallback)
{
	AIChatResult result = {};
	std::string missingField;
	if (!HasRequiredSettings(settings, missingField)) {
		result.error = "AI settings missing: " + missingField;
		return result;
	}

	if (settings.protocolType == AIProtocolType::Claude) {
		return ExecuteChatWithToolsClaude(contextMessages, settings, toolCallback, streamCallback);
	}
	if (settings.protocolType == AIProtocolType::Gemini) {
		return ExecuteChatWithToolsGemini(contextMessages, settings, toolCallback, streamCallback);
	}

	const std::string endpoint = BuildEndpoint(settings.baseUrl);
	const std::string headers = BuildOpenAIHeaders(settings);
	const uint64_t traceId = GetCurrentAIPerfTraceId();

	nlohmann::json requestMessages = nlohmann::json::array();
	requestMessages.push_back({
		{"role", "system"},
		{"content", LocalToUtf8(BuildChatSystemPrompt(settings))}
	});
	for (const AIChatMessage& msg : contextMessages) {
		const std::string role = ToLowerAsciiCopy(Trim(msg.role));
		if (role != "system" && role != "user" && role != "assistant") {
			continue;
		}
		requestMessages.push_back({
			{"role", role},
			{"content", LocalToUtf8(msg.content)}
		});
	}

	const nlohmann::json tools = BuildChatToolDefinitions();
	const int maxToolRounds = GetMaxToolRounds(settings);

	for (int round = 0; round < maxToolRounds; ++round) {
		nlohmann::json requestBody;
		requestBody["model"] = LocalToUtf8(settings.model);
		requestBody["temperature"] = settings.temperature;
		requestBody["stream"] = true;
		requestBody["messages"] = requestMessages;
		requestBody["tools"] = tools;
		requestBody["tool_choice"] = "auto";

		std::string requestBodyText;
		try {
			requestBodyText = requestBody.dump();
		}
		catch (const std::exception& ex) {
			result.error = std::string("Failed to build AI chat request JSON: ") + ex.what();
			return result;
		}

		ChatStreamParseState streamState;
		const auto networkStart = PerfClock::now();
		const auto [responseBody, statusCode] =
			PerformPostRequestStreamingWithRetry(
				endpoint,
				requestBodyText,
				[&streamState, &streamCallback](const std::string& chunk) -> bool {
					return ConsumeStreamChunk(chunk, streamState, streamCallback);
				},
				headers,
				settings.timeoutMs,
				false,
				false,
				"openai-chat");
		LogAIPerfCost(
			traceId,
			"AIService.ExecuteTask.network_total",
			ElapsedMs(networkStart),
			"http=" + std::to_string(statusCode) + " endpoint=" + endpoint);
		result.httpStatus = statusCode;
		if (statusCode < 200 || statusCode >= 300) {
			result.error = std::format("HTTP {}: {}", statusCode, TruncateForLog(responseBody));
			return result;
		}

		if (!FlushStreamParseState(streamState, streamCallback)) {
			result.error = streamState.parseError.empty() ? "Failed to parse AI streaming response" : streamState.parseError;
			return result;
		}

		nlohmann::json message;
		if (streamState.sawDataEvent) {
			if (!streamState.parseError.empty()) {
				result.error = streamState.parseError;
				return result;
			}
			message = BuildAssistantMessageFromStreamState(streamState);
		}
		else {
			nlohmann::json parsed;
			try {
				parsed = nlohmann::json::parse(responseBody);
			}
			catch (const std::exception& ex) {
				result.error = std::string("Failed to parse AI response: ") + ex.what();
				return result;
			}

			std::string parseError;
			if (!ExtractChatResponseMessage(parsed, message, parseError)) {
				if (parsed.contains("error") && parsed["error"].contains("message") && parsed["error"]["message"].is_string()) {
					result.error = Utf8ToLocal(parsed["error"]["message"].get<std::string>());
				}
				else {
					result.error = parseError.empty() ? "AI response parse failed" : parseError;
				}
				return result;
			}
		}

		// Tool-call path.
		if (message.contains("tool_calls") && message["tool_calls"].is_array() && !message["tool_calls"].empty()) {
			requestMessages.push_back(message);

			for (const auto& toolCall : message["tool_calls"]) {
				std::string callId;
				std::string toolName;
				std::string argsUtf8;
				if (toolCall.contains("id") && toolCall["id"].is_string()) {
					callId = toolCall["id"].get<std::string>();
				}
				if (toolCall.contains("function") && toolCall["function"].is_object()) {
					const auto& fn = toolCall["function"];
					if (fn.contains("name") && fn["name"].is_string()) {
						toolName = fn["name"].get<std::string>();
					}
					if (fn.contains("arguments") && fn["arguments"].is_string()) {
						argsUtf8 = fn["arguments"].get<std::string>();
					}
				}

				if (callId.empty()) {
					callId = std::format("call_auto_round{}_{}", round + 1, result.toolEvents.size() + 1);
				}

				bool toolOk = false;
				std::string toolResultLocal;
				if (toolCallback) {
					toolResultLocal = toolCallback(toolName, argsUtf8, toolOk);
				}
				else {
					toolResultLocal = R"({"ok":false,"error":"tool callback not set"})";
					toolOk = false;
				}
				const CompactToolResultPayload compactPayload = BuildCompactToolResultPayload(toolName, toolResultLocal);

				AIChatToolEvent evt = {};
				evt.name = toolName;
				evt.argumentsJson = Utf8ToLocal(argsUtf8);
				evt.resultJson = toolResultLocal;
				evt.ok = toolOk;
				result.toolEvents.push_back(std::move(evt));

				requestMessages.push_back({
					{"role", "tool"},
					{"tool_call_id", callId},
					{"name", toolName},
					{"content", compactPayload.textUtf8}
				});
			}
			continue;
		}

		// Final assistant content path.
		std::string mergedUtf8 = MergeMessageContentUtf8(message);
		if (!streamState.sawDataEvent && streamCallback && !mergedUtf8.empty()) {
			streamCallback(Utf8ToLocal(mergedUtf8));
		}
		if (mergedUtf8.empty()) {
			result.error = "AI response content is empty";
			return result;
		}

		result.ok = true;
		result.content = Utf8ToLocal(mergedUtf8);
		return result;
	}

	result.error = BuildToolRoundsExceededError(maxToolRounds, result.toolEvents);
	return result;
}

std::string AIService::BuildPublicToolCatalogJson()
{
	return BuildPublicToolCatalog().dump();
}

std::string AIService::NormalizeModelOutputToCode(const std::string& modelText)
{
	return RemoveCodeFence(modelText);
}

std::string AIService::Trim(const std::string& text)
{
	size_t begin = 0;
	size_t end = text.size();
	while (begin < end && std::isspace(static_cast<unsigned char>(text[begin])) != 0) {
		++begin;
	}
	while (end > begin && std::isspace(static_cast<unsigned char>(text[end - 1])) != 0) {
		--end;
	}
	return text.substr(begin, end - begin);
}

std::string AIService::BuildEndpoint(const std::string& baseUrl)
{
	std::string url = Trim(baseUrl);
	while (!url.empty() && url.back() == '/') {
		url.pop_back();
	}
	if (EndsWithInsensitive(url, "/chat/completions")) {
		return url;
	}
	if (EndsWithInsensitive(url, "/v1")) {
		return url + "/chat/completions";
	}
	return url + "/v1/chat/completions";
}

std::string AIService::BuildSystemPrompt(AITaskKind kind, const AISettings& settings)
{
	const std::string projectType = DetectProjectTypeText();
	std::string prompt =
		"你是一个易语言代码助手。\n"
		"当前项目类型：" + projectType + "\n"
		"规则：\n"
		"1) 代码注释使用单引号（'）。\n"
		"2) 必须遵循易语言语法与排版（如 .版本 2、.子程序 等）。\n"
		"3) 除非用户明确要求，否则只输出最终代码或文本，不要附加解释。\n"
		"4) 不要输出 Markdown 标题或其他包装。\n\n"
		"易语言格式示例：\n"
		".版本 2\n"
		".子程序 demo, 整数型\n"
		"返回 (0)\n\n";

	prompt += R"AL_REF(参考函数（完整示例，仅用于格式与风格参考，不要照抄无关变量）：
.版本 2

.子程序 FindWindowWithContainReturnMainList, 整数型, 公开, 函数注释：寻找所有符合条件的顶级窗口句柄，返回找到的数量
.参数 mainTitle, 文本型, , 参数；主窗口标题包含的文本
.参数 mainClass, 文本型, , 参数；主窗口类名包含的文本
.参数 childTitle, 文本型, , 参数；子窗口标题包含的文本
.参数 childClass, 文本型, , 参数；子窗口类名包含的文本
.参数 childHasChild, 逻辑型, 可空, 参数；可空；找到的子窗口需要包含子窗口(真)还是不包含(假)
.参数 mainHwndArray, 整数型, 参考 数组, 参数；参考（传址）；数组；用于存放找到的所有主窗口句柄的数组变量
.局部变量 i, 整数型, , , 这些都是变量
.局部变量 class, 文本型
.局部变量 text, 文本型
.局部变量 mid, 整数型, , "0", 这是个数组
.局部变量 x, 整数型
.局部变量 mainOK, 逻辑型
.局部变量 mainTitleOK, 逻辑型
.局部变量 mainClassOK, 逻辑型
.局部变量 childTitleOK, 逻辑型
.局部变量 childClassOK, 逻辑型
.局部变量 childOK, 逻辑型
.局部变量 hasChild, 逻辑型

' 1. 初始化结果数组
清除数组 (mainHwndArray)

' 2. 获取当前所有顶级窗口
清除数组 (m_hwnd_list)
EnumWindows (到整数 (&枚举窗口过程), 0)

' 3. 遍历顶级窗口
.计次循环首 (取数组成员数 (m_hwnd_list), i)
    text ＝ 窗口_取标题 (m_hwnd_list [i])
    class ＝ 窗口_取类名 (m_hwnd_list [i])

    ' 匹配主窗口条件 (空字符串视为匹配成功)
    mainTitleOK ＝ 选择 (mainTitle ＝ "", 真, IsContains (text, mainTitle))
    mainClassOK ＝ 选择 (mainClass ＝ "", 真, IsContains (class, mainClass))

    mainOK ＝ mainTitleOK 且 mainClassOK

    .如果真 (mainOK)
        ' 4. 如果主窗口匹配，枚举其所有子窗口进行深度检查
        清除数组 (mid)
        窗口_枚举所有子窗口 (m_hwnd_list [i], mid, )

        .计次循环首 (取数组成员数 (mid), x)
            text ＝ 窗口_取标题 (mid [x])
            class ＝ 窗口_取类名 (mid [x])

            ' 匹配子窗口条件
            childTitleOK ＝ 选择 (childTitle ＝ "", 真, IsContains (text, childTitle))
            childClassOK ＝ 选择 (childClass ＝ "", 真, IsContains (class, childClass))
            childOK ＝ childTitleOK 且 childClassOK

            ' 检查子窗口是否含有孙窗口
            hasChild ＝ hasChildWindow (mid [x])

            ' 综合判断子窗口是否符合要求
            .如果 (childHasChild)
                .如果真 (childOK 且 hasChild)
                    加入成员 (mainHwndArray, m_hwnd_list [i])
                    跳出循环 ()  ' 只要找到一个符合条件的子窗口，该主窗口就合格，跳出子窗口循环
                .如果真结束

            .否则
                .如果真 (childOK 且 hasChild ＝ 假)
                    加入成员 (mainHwndArray, m_hwnd_list [i])
                    跳出循环 ()
                .如果真结束

            .如果结束

        .计次循环尾 ()
    .如果真结束

.计次循环尾 ()

' 5. 返回找到的总数
返回 (取数组成员数 (mainHwndArray))

)AL_REF";

	switch (kind)
	{
	case AITaskKind::OptimizeFunction:
		prompt += "任务：优化给定函数代码，在保持行为等价的前提下提升可读性与健壮性。只返回完整可替换的函数代码。";
		break;
	case AITaskKind::AddCommentsToFunction:
		prompt += "任务：为给定函数添加合适注释（函数说明与关键行注释），不得改变原逻辑。只返回完整可替换的函数代码。";
		break;
	case AITaskKind::TranslateFunctionAndVariables:
		prompt +=
			"任务：将函数名、参数名、局部变量名翻译或重命名为英文 lowerCamelCase（首字母小写），并保持逻辑不变。\n"
			"禁止翻译或修改任何以 '.' 开头的易语言系统指令/关键字（例如：.版本/.子程序/.参数/.局部变量/.如果/.否则/.返回 等）。\n"
			"只允许修改标识符（函数名/参数名/局部变量名），系统指令与语句结构必须保持不变。\n"
			"只返回完整可替换的函数代码。";
		break;
	case AITaskKind::TranslateText:
		prompt += "任务：翻译用户提供的文本。只返回翻译后的纯文本，不要附加解释。";
		break;
	case AITaskKind::AddByCurrentPageType:
		prompt +=
			"任务：根据‘当前页类型 + 用户需求 + 当前页伪代码’生成一段可直接追加的代码片段。\n\n"
			"额外要求：\n"
			"1) 只输出要追加的代码，禁止重复整个页面。\n"
			"2) 不要输出解释或 Markdown 包装。\n"
			"3) 不要重复输出 .版本 行。\n"
			"4) 必须遵守当前页面已有结构和书写风格。\n"
			"5) 生成结果必须能直接粘贴到当前页面末尾。\n\n"
			"输出时保持原有换行与缩进，不要把多行代码压成一行。";
		break;
	default:
		break;
	}

	const std::string extraPrompt = Trim(settings.extraSystemPrompt);
	if (!extraPrompt.empty()) {
		prompt += "\n\n用户额外系统提示：\n";
		prompt += extraPrompt;
	}

	return prompt;
}

