#include "AIService.h"

#include <algorithm>
#include <cctype>
#include <format>

#include "..\\thirdparty\\json.hpp"

#include "ConfigManager.h"
#include "WinINetUtil.h"

namespace {
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
} // namespace

bool AIService::LoadSettings(ConfigManager& config, AISettings& outSettings)
{
	outSettings = {};
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

std::string AIService::BuildTaskDisplayName(AITaskKind kind)
{
	switch (kind)
	{
	case AITaskKind::OptimizeFunction:
		return "AI Optimize Function";
	case AITaskKind::AddCommentsToFunction:
		return "AI Add Function Comments";
	case AITaskKind::TranslateFunctionAndVariables:
		return "AI Translate Function And Variables";
	case AITaskKind::TranslateText:
		return "AI Translate Text";
	case AITaskKind::CompleteApiDeclarations:
		return "AI Complete API Declarations";
	default:
		return "AI Task";
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

	nlohmann::json requestBody;
	requestBody["model"] = settings.model;
	requestBody["temperature"] = settings.temperature;
	requestBody["stream"] = false;
	requestBody["messages"] = nlohmann::json::array({
		{
			{"role", "system"},
			{"content", BuildSystemPrompt(kind, settings)}
		},
		{
			{"role", "user"},
			{"content", inputText}
		}
	});

	const std::string endpoint = BuildEndpoint(settings.baseUrl);
	const std::string headers =
		"Content-Type: application/json\r\n"
		"Authorization: Bearer " + settings.apiKey + "\r\n";

	const auto [responseBody, statusCode] =
		PerformPostRequest(endpoint, requestBody.dump(), headers, settings.timeoutMs, false, false);
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
					result.content = choice["message"]["content"].get<std::string>();
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
						result.content = merged;
						return result;
					}
				}
			}
		}

		if (parsed.contains("error") && parsed["error"].contains("message") && parsed["error"]["message"].is_string()) {
			result.error = parsed["error"]["message"].get<std::string>();
			return result;
		}

		result.error = "AI response does not match expected chat/completions schema";
	}
	catch (const std::exception& ex) {
		result.error = std::string("Failed to parse AI response: ") + ex.what();
	}

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
		"You are a coding assistant for E-language (Yi language), a Chinese programming language.\n"
		"Rules:\n"
		"1) Use single quote (') for inline comments in code.\n"
		"2) Follow E-language syntax and keep '.version 2' style layout.\n"
		"3) Unless explicitly requested, output only final code/text, no extra explanation.\n"
		"4) Do not output markdown headings or extra wrappers.\n\n"
		"E-language layout example:\n"
		".version 2\n"
		".subroutine demo, integer\n"
		"return (0)\n\n";

	switch (kind)
	{
	case AITaskKind::OptimizeFunction:
		prompt += "Task: optimize the provided current function code for readability and robustness while keeping behavior equivalent. Return only complete function code.";
		break;
	case AITaskKind::AddCommentsToFunction:
		prompt += "Task: add appropriate comments to the provided function (function comment and key-line comments) without changing logic. Return only complete function code.";
		break;
	case AITaskKind::TranslateFunctionAndVariables:
		prompt += "Task: rename function name, parameter names, and local variable names to English using lowerCamelCase (first letter lowercase) while preserving logic. Return only complete function code.";
		break;
	case AITaskKind::TranslateText:
		prompt += "Task: translate user-provided text. Return only translated plain text with no extra explanation.";
		break;
	case AITaskKind::CompleteApiDeclarations:
		prompt +=
			"Task: based on the provided function, complete potentially missing .DLL command / .parameter / .data type / .member declarations.\n"
			"Return declarations only, do not return function implementation.\n"
			"Declaration example:\n"
			".DLL command GetDC, integer, \"user32\", \"GetDC\", public\n"
			".parameter hwnd, integer\n"
			".data type BITMAPFILEHEADER\n"
			".member bfType, short integer";
		break;
	default:
		break;
	}

	const std::string extraPrompt = Trim(settings.extraSystemPrompt);
	if (!extraPrompt.empty()) {
		prompt += "\n\nUser extra system prompt:\n";
		prompt += extraPrompt;
	}

	return prompt;
}
