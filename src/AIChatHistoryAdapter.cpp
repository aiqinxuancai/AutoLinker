#include "AIChatHistoryAdapter.h"
#include "StaticJson.h"

#include <string>
#include <utility>

namespace {
using Json = nlohmann::json;

std::string ReadString(const Json& item, const char* key)
{
	const auto it = item.find(key);
	return it != item.end() && it->is_string() ? it->get<std::string>() : std::string();
}

Json ConvertContent(const Json& content)
{
	if (content.is_string()) {
		return content;
	}
	if (!content.is_array()) {
		return "";
	}
	Json parts = Json::array();
	std::string text;
	bool hasImage = false;
	for (const auto& part : content) {
		const std::string type = ReadString(part, "type");
		if (type == "output_text" || type == "input_text" || type == "text" || type == "refusal") {
			const std::string value = ReadString(part, type == "refusal" ? "refusal" : "text");
			text += value;
			parts.push_back({{"type", "text"}, {"text", value}});
		}
		else if (type == "input_image" && !ReadString(part, "image_url").empty()) {
			Json image = {{"url", part["image_url"]}};
			if (!ReadString(part, "detail").empty()) {
				image["detail"] = part["detail"];
			}
			parts.push_back({{"type", "image_url"}, {"image_url", std::move(image)}});
			hasImage = true;
		}
	}
	return hasImage ? parts : Json(text);
}

bool IsUserImageMessage(const Json& message)
{
	if (ReadString(message, "role") != "user" || !message.contains("content") || !message["content"].is_array()) {
		return false;
	}
	for (const auto& part : message["content"]) {
		if (ReadString(part, "type") == "image_url") {
			return true;
		}
	}
	return false;
}

// Responses 工具图片可能夹在两个结果之间；Chat 要求整批结果紧跟助手调用。
void MoveToolImagesAfterBatch(Json& messages)
{
	Json ordered = Json::array();
	for (size_t i = 0; i < messages.size(); ++i) {
		const auto& message = messages[i];
		ordered.push_back(message);
		if (ReadString(message, "role") != "assistant" || !message.contains("tool_calls")) {
			continue;
		}
		Json images = Json::array();
		while (i + 1 < messages.size()) {
			const auto& next = messages[i + 1];
			if (ReadString(next, "role") == "tool") {
				ordered.push_back(next);
			}
			else if (IsUserImageMessage(next)) {
				images.push_back(next);
			}
			else {
				break;
			}
			++i;
		}
		for (auto& image : images) {
			ordered.push_back(std::move(image));
		}
	}
	messages = std::move(ordered);
}

}

void AIChatHistoryAdapter::ConvertResponsesItemsToChat(Json& messages, bool includeReasoningContent)
{
	if (!messages.is_array()) {
		return;
	}
	Json converted = Json::array();
	Json assistant;
	const auto ensureAssistant = [&]() {
		if (assistant.is_null()) {
			assistant = AUTOLINKER_STATIC_JSON(R"json({"role":"assistant","content":""})json");
			if (includeReasoningContent) {
				// GPT 的加密推理无法转换；只为目标协议提供空值，绝不伪造思考正文。
				assistant["reasoning_content"] = "";
			}
		}
	};
	const auto flush = [&]() {
		if (!assistant.is_null()) {
			converted.push_back(std::move(assistant));
			assistant = nullptr;
		}
	};
	for (const auto& item : messages) {
		const std::string type = ReadString(item, "type");
		if (type == "reasoning" || type == "previous_response_ref") {
			continue;
		}
		if (type == "function_call") {
			const std::string callId = ReadString(item, "call_id");
			const std::string name = ReadString(item, "name");
			if (callId.empty() || name.empty()) {
				continue;
			}
			ensureAssistant();
			if (!assistant.contains("tool_calls")) {
				assistant["tool_calls"] = Json::array();
			}
			assistant["tool_calls"].push_back({{"id", callId}, {"type", "function"},
				{"function", {{"name", name}, {"arguments", ReadString(item, "arguments")}}}});
			continue;
		}
		if (type == "message") {
			const std::string role = ReadString(item, "role");
			const Json content = ConvertContent(item.value("content", Json()));
			if (role == "assistant" && content.is_string()) {
				ensureAssistant();
				assistant["content"] = assistant["content"].get<std::string>() + content.get<std::string>();
			}
			else {
				flush();
				converted.push_back({{"role", role}, {"content", content}});
			}
			continue;
		}
		flush();
		if (type == "function_call_output") {
			const std::string callId = ReadString(item, "call_id");
			if (!callId.empty()) {
				converted.push_back({{"role", "tool"}, {"tool_call_id", callId},
					{"content", ConvertContent(item.value("output", Json()))}});
			}
		}
		else {
			// 原生 DeepSeek / Chat 消息及其 reasoning_content 保持原样。
			converted.push_back(item);
		}
	}
	flush();
	MoveToolImagesAfterBatch(converted);
	messages = std::move(converted);
}
