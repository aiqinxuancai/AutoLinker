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
}

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
			outSettings.timeoutMs = std::max(1000, std::stoi(timeoutValue));
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
		return "AI优化函数";
	case AITaskKind::AddCommentsToFunction:
		return "AI添加函数注释";
	case AITaskKind::TranslateFunctionAndVariables:
		return "AI翻译函数与变量名";
	case AITaskKind::TranslateText:
		return "AI翻译文本";
	case AITaskKind::CompleteApiDeclarations:
		return "AI补全API声明";
	default:
		return "AI任务";
	}
}

AIResult AIService::ExecuteTask(AITaskKind kind, const std::string& inputText, const AISettings& settings)
{
	AIResult result = {};
	std::string missingField;
	if (!HasRequiredSettings(settings, missingField)) {
		result.error = "AI配置缺失: " + missingField;
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

		if (parsed.contains("error") && parsed["error"].contains("message")) {
			result.error = parsed["error"]["message"].get<std::string>();
			return result;
		}

		result.error = "AI响应格式不符合 chat/completions 预期";
	}
	catch (const std::exception& ex) {
		result.error = std::string("解析AI响应失败: ") + ex.what();
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
		"你是中文编程语言“易语言”代码助手。\n"
		"严格遵守以下规则：\n"
		"1) 注释使用单引号 '。\n"
		"2) 代码结构遵循 .版本 2 和易语言语法。\n"
		"3) 如非特别要求，仅返回结果代码或结果文本，不要解释。\n"
		"4) 不要输出 Markdown 标题或多余说明。\n\n"
		"易语言布局示例：\n"
		".版本 2\n"
		".子程序 示例, 整数型\n"
		"返回 (0)\n\n";

	switch (kind)
	{
	case AITaskKind::OptimizeFunction:
		prompt += "任务：优化用户提供的当前函数代码，可提升可读性与健壮性，但必须保持功能等价。仅输出完整函数代码。";
		break;
	case AITaskKind::AddCommentsToFunction:
		prompt += "任务：为用户提供的函数添加适量注释（函数注释、关键行注释），不要改变代码逻辑。仅输出完整函数代码。";
		break;
	case AITaskKind::TranslateFunctionAndVariables:
		prompt += "任务：将函数名、参数名、局部变量名改为英文，使用 lowerCamelCase（首字母小写），逻辑不变。仅输出完整函数代码。";
		break;
	case AITaskKind::TranslateText:
		prompt += "任务：翻译用户提供的文本。仅输出翻译结果纯文本，不要附加解释。";
		break;
	case AITaskKind::CompleteApiDeclarations:
		prompt +=
			"任务：基于用户提供函数，补全可能缺失的 .DLL命令 / .参数 / .数据类型 / .成员 声明。\n"
			"只返回声明代码块，不返回函数实现。\n"
			"声明示例：\n"
			".DLL命令 GetDC, 整数型, \"user32\", \"GetDC\", 公开\n"
			".参数 hwnd, 整数型\n"
			".数据类型 BITMAPFILEHEADER\n"
			".成员 bfType, 短整数型";
		break;
	default:
		break;
	}

	const std::string extraPrompt = Trim(settings.extraSystemPrompt);
	if (!extraPrompt.empty()) {
		prompt += "\n\n用户追加系统提示：\n";
		prompt += extraPrompt;
	}

	return prompt;
}
