#include "AIService.h"

#include <algorithm>
#include <cctype>
#include <format>
#include <Windows.h>

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

nlohmann::json BuildChatToolDefinitions()
{
	nlohmann::json tools = nlohmann::json::array();
	tools.push_back({
		{"type", "function"},
		{"function", {
			{"name", "get_current_page_code"},
			{"description", "Get complete source code of the current IDE page."},
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
		"你可以通过工具获取当前页代码，或者请求用户在本地可编辑对话框中提供代码。\n"
		"规则：\n"
		"1) 需要源码时优先调用 get_current_page_code，不要臆造现有代码。\n"
		"2) 需要用户确认/修订代码时调用 request_code_edit。\n"
		"3) 工具返回失败或取消时，给出下一步建议，不要编造工具结果。\n"
		"4) 除非用户要求解释，否则尽量给直接可执行结论。\n";

	const std::string extraPrompt = AIService::Trim(settings.extraSystemPrompt);
	if (!extraPrompt.empty()) {
		prompt += "\n附加系统提示：\n";
		prompt += extraPrompt;
	}
	return prompt;
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

	const std::string modelUtf8 = LocalToUtf8(settings.model);
	const std::string systemPromptUtf8 = LocalToUtf8(BuildSystemPrompt(kind, settings));
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
	const std::string headers =
		"Content-Type: application/json\r\n"
		"Authorization: Bearer " + settings.apiKey + "\r\n";

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
	const std::function<std::string(const std::string& toolName, const std::string& argumentsJson, bool& outOk)>& toolCallback)
{
	AIChatResult result = {};
	std::string missingField;
	if (!HasRequiredSettings(settings, missingField)) {
		result.error = "AI settings missing: " + missingField;
		return result;
	}

	const std::string endpoint = BuildEndpoint(settings.baseUrl);
	const std::string headers =
		"Content-Type: application/json\r\n"
		"Authorization: Bearer " + settings.apiKey + "\r\n";

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
		requestBody["stream"] = false;
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

		const auto [responseBody, statusCode] =
			PerformPostRequest(endpoint, requestBodyText, headers, settings.timeoutMs, false, false);
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
			result.error = std::string("Failed to parse AI response: ") + ex.what();
			return result;
		}

		if (!parsed.contains("choices") || !parsed["choices"].is_array() || parsed["choices"].empty()) {
			result.error = "AI response choices is empty";
			return result;
		}

		const nlohmann::json& choice = parsed["choices"][0];
		if (!choice.contains("message") || !choice["message"].is_object()) {
			result.error = "AI response message missing";
			return result;
		}
		const nlohmann::json& message = choice["message"];

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
		std::string mergedUtf8;
		if (message.contains("content") && message["content"].is_string()) {
			mergedUtf8 = message["content"].get<std::string>();
		}
		else if (message.contains("content") && message["content"].is_array()) {
			for (const auto& item : message["content"]) {
				if (item.is_string()) {
					mergedUtf8 += item.get<std::string>();
				}
				else if (item.is_object() && item.contains("text") && item["text"].is_string()) {
					mergedUtf8 += item["text"].get<std::string>();
				}
			}
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

