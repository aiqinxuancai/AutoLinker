#include "AIService.h"

#include <algorithm>
#include <cctype>
#include <format>
#include <Windows.h>

#include "..\\thirdparty\\json.hpp"

#include "ConfigManager.h"
#include "Global.h"
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

std::string TruncateForLog(const std::string& text, size_t maxLen = 240)
{
	if (text.size() <= maxLen) {
		return text;
	}
	return text.substr(0, maxLen) + "...";
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

nlohmann::json BuildChatToolDefinitions()
{
	nlohmann::json tools = nlohmann::json::array();
	tools.push_back({
		{"type", "function"},
		{"function", {
			{"name", "get_current_page_code"},
			{"description", "Get complete source code of the current IDE page, together with current page name and page type."},
			{"parameters", {
				{"type", "object"},
				{"properties", nlohmann::json::object()},
				{"additionalProperties", false}
			}}
		}}
	});
	tools.push_back({
		{"type", "function"},
		{"function", {
			{"name", "get_current_page_info"},
			{"description", "Get current IDE page name, page type and the trace/source used to resolve that page name."},
			{"parameters", {
				{"type", "object"},
				{"properties", nlohmann::json::object()},
				{"additionalProperties", false}
			}}
		}}
	});
	tools.push_back({
		{"type", "function"},
		{"function", {
			{"name", "list_program_items"},
			{"description", "List program tree items such as assemblies, class modules, global variables, user-defined types, DLL commands, forms and resources. Can optionally include code for each item. Important: code returned by program-tree lookup is only pseudo-code reference and may differ from the normal IDE page structure."},
			{"parameters", {
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
		}}
	});
	tools.push_back({
		{"type", "function"},
		{"function", {
			{"name", "get_program_item_code"},
			{"description", "Get code of a program tree item by exact name, optionally constrained by kind. Important: this may switch the IDE current page as part of native retrieval, and returned code is only pseudo-code reference and may differ from the normal IDE page structure."},
			{"parameters", {
				{"type", "object"},
				{"properties", {
					{"name", {{"type", "string"}}},
					{"kind", {{"type", "string"}, {"description", "Optional kind filter: assembly, class_module, global_var, user_data_type, dll_command, form, const_resource, picture_resource, sound_resource."}}}
				}},
				{"required", nlohmann::json::array({"name"})},
				{"additionalProperties", false}
			}}
		}}
	});
	tools.push_back({
		{"type", "function"},
		{"function", {
			{"name", "switch_to_program_item_page"},
			{"description", "Switch/open a program tree page by exact name, optionally constrained by kind. This will change the IDE current page and only activates that page; it does not fetch code."},
			{"parameters", {
				{"type", "object"},
				{"properties", {
					{"name", {{"type", "string"}}},
					{"kind", {{"type", "string"}, {"description", "Optional kind filter: assembly, class_module, global_var, user_data_type, dll_command, form, const_resource, picture_resource, sound_resource."}}}
				}},
				{"required", nlohmann::json::array({"name"})},
				{"additionalProperties", false}
			}}
		}}
	});
	tools.push_back({
		{"type", "function"},
		{"function", {
			{"name", "search_project_keyword"},
			{"description", "Search a keyword across the project using IDE global search and return matched page names, line numbers and a jump_token for each result. Search-based line text and follow-up code lookup should be treated as pseudo-code reference, not exact normal IDE page structure."},
			{"parameters", {
				{"type", "object"},
				{"properties", {
					{"keyword", {{"type", "string"}}},
					{"limit", {{"type", "integer"}, {"minimum", 1}, {"maximum", 200}}}
				}},
				{"required", nlohmann::json::array({"keyword"})},
				{"additionalProperties", false}
			}}
		}}
	});
	tools.push_back({
		{"type", "function"},
		{"function", {
			{"name", "jump_to_search_result"},
			{"description", "Jump to one specific search result returned by search_project_keyword using that row's jump_token. This will change the IDE current page and caret position."},
			{"parameters", {
				{"type", "object"},
				{"properties", {
					{"jump_token", {{"type", "string"}}}
				}},
				{"required", nlohmann::json::array({"jump_token"})},
				{"additionalProperties", false}
			}}
		}}
	});
	tools.push_back({
		{"type", "function"},
		{"function", {
			{"name", "request_code_edit"},
			{"description", "Open local editable code dialog and return user confirmed code."},
			{"parameters", {
				{"type", "object"},
				{"properties", {
					{"title", {{"type", "string"}}},
					{"initial_code", {{"type", "string"}}},
					{"hint", {{"type", "string"}}}
				}},
				{"required", nlohmann::json::array({"title", "initial_code"})},
				{"additionalProperties", false}
			}}
		}}
	});
	return tools;
}

std::string BuildChatSystemPrompt(const AISettings& settings)
{
	std::string prompt =
		"你是 AutoLinker 的易语言开发助手。\n"
		"你可以通过工具获取当前页代码、枚举程序树中的程序集/类/全局变量等、按名称抓整页代码，以及执行项目级关键词搜索。\n"
		"规则：\n"
		"1) 需要源码时优先调用 get_current_page_code，不要臆造现有代码；该工具会同时返回当前页名称和页类型。\n"
		"2) 只需要知道当前页是谁而不需要全文代码时，调用 get_current_page_info。\n"
		"3) 需要枚举项目结构时优先调用 list_program_items；需要某个程序集/类整页代码时调用 get_program_item_code；只需要切换页面时调用 switch_to_program_item_page。\n"
		"4) 需要项目内关键词定位时先调用 search_project_keyword 查看结果；若要精确跳到其中某一条，使用该条结果返回的 jump_token 调用 jump_to_search_result。\n"
		"5) get_program_item_code、switch_to_program_item_page、jump_to_search_result 都可能触发 IDE 当前页面变更；调用前要意识到这是有副作用的，不要把调用前后的当前页混为一谈。\n"
		"6) 通过搜索结果或程序树按名称拿到的代码，不保证与 IDE 正常编辑页结构一致，只能作为伪代码参考；分析和修改建议时必须明确这一点。\n"
		"7) 需要用户确认/修订代码时调用 request_code_edit。\n"
		"8) 工具返回失败或取消时，给出下一步建议，不要编造工具结果。\n"
		"9) 除非用户要求解释，否则尽量给直接可执行结论。\n";

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

	const auto [responseBody, statusCode] = PerformPostRequest(
		endpoint,
		requestBody.dump(),
		BuildClaudeHeaders(settings),
		settings.timeoutMs,
		false,
		false);
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

	const auto [responseBody, statusCode] = PerformPostRequest(
		endpoint,
		requestBody.dump(),
		BuildJsonHeadersOnly(),
		settings.timeoutMs,
		false,
		false);
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

	constexpr int kMaxToolRounds = 8;
	for (int round = 0; round < kMaxToolRounds; ++round) {
		nlohmann::json requestBody;
		requestBody["model"] = LocalToUtf8(settings.model);
		requestBody["max_tokens"] = 4096;
		requestBody["temperature"] = settings.temperature;
		requestBody["system"] = systemUtf8;
		requestBody["messages"] = messages;
		requestBody["tools"] = tools;
		requestBody["tool_choice"] = { {"type", "auto"} };
		requestBody["stream"] = false;

		const auto [responseBody, statusCode] = PerformPostRequest(
			endpoint,
			requestBody.dump(),
			BuildClaudeHeaders(settings),
			settings.timeoutMs,
			false,
			false);
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
						{"content", LocalToUtf8(toolResultLocal)}
					}
				})}
			});
		}
	}

	result.error = "tool call rounds exceeded limit";
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

	constexpr int kMaxToolRounds = 8;
	for (int round = 0; round < kMaxToolRounds; ++round) {
		nlohmann::json requestBody;
		requestBody["system_instruction"] = {
			{"parts", nlohmann::json::array({ {{"text", systemUtf8}} })}
		};
		requestBody["generationConfig"] = { {"temperature", settings.temperature} };
		requestBody["contents"] = contents;
		requestBody["tools"] = tools;

		const auto [responseBody, statusCode] = PerformPostRequest(
			endpoint,
			requestBody.dump(),
			BuildJsonHeadersOnly(),
			settings.timeoutMs,
			false,
			false);
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

			AIChatToolEvent evt = {};
			evt.name = call.name;
			evt.argumentsJson = Utf8ToLocal(call.argumentsUtf8);
			evt.resultJson = toolResultLocal;
			evt.ok = toolOk;
			result.toolEvents.push_back(std::move(evt));

			nlohmann::json toolResultNode;
			try {
				toolResultNode = nlohmann::json::parse(LocalToUtf8(toolResultLocal));
			}
			catch (...) {
				toolResultNode = {
					{"ok", toolOk},
					{"text", LocalToUtf8(toolResultLocal)}
				};
			}

			contents.push_back({
				{"role", "user"},
				{"parts", nlohmann::json::array({
					{
						{"functionResponse", {
							{"name", call.name},
							{"response", toolResultNode}
						}}
					}
				})}
			});
		}
	}

	result.error = "tool call rounds exceeded limit";
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

	return true;
}

void AIService::SaveSettings(ConfigManager& config, const AISettings& settings)
{
	config.setValue("ai.protocol_type", ProtocolTypeToString(settings.protocolType));
	config.setValue("ai.base_url", settings.baseUrl);
	config.setValue("ai.api_key", settings.apiKey);
	config.setValue("ai.model", settings.model);
	config.setValue("ai.system_prompt_extra", settings.extraSystemPrompt);
	config.setValue("ai.timeout_ms", std::to_string(settings.timeoutMs));
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
		PerformPostRequest(endpoint, requestBodyText, headers, settings.timeoutMs, false, false);
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
	constexpr int kMaxToolRounds = 8;

	for (int round = 0; round < kMaxToolRounds; ++round) {
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
			PerformPostRequestStreaming(
				endpoint,
				requestBodyText,
				[&streamState, &streamCallback](const std::string& chunk) -> bool {
					return ConsumeStreamChunk(chunk, streamState, streamCallback);
				},
				headers,
				settings.timeoutMs,
				false,
				false);
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
					{"content", LocalToUtf8(toolResultLocal)}
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

	result.error = "tool call rounds exceeded limit";
	return result;
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
	std::string prompt =
		"你是一个易语言代码助手。\n"
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
			"任务：根据“当前页类型 + 用户需求 + 当前页完整代码”，生成“仅新增”的代码片段。\n"
			"严格要求：\n"
			"1) 只返回要新增的代码，禁止返回整页代码。\n"
			"2) 不要输出解释、不要输出 Markdown。\n"
			"3) 不要输出 .版本 行。\n"
			"4) 避免与当前页已有定义重复命名。\n"
			"5) 输出内容需可直接粘贴到当前页底部。\n";
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

