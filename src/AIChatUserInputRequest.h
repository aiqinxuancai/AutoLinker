#pragma once

#include <string>
#include <vector>

// AI 计划模式中的单个可选答案。
struct AIChatUserInputOption {
	std::string labelUtf8;
	std::string descriptionUtf8;
};

// AI 计划模式中的单个结构化问题。
struct AIChatUserInputQuestion {
	std::string id;
	std::string headerUtf8;
	std::string questionUtf8;
	std::vector<AIChatUserInputOption> options;
};

// 用户对单个结构化问题的回答。
struct AIChatUserInputAnswer {
	std::string questionId;
	std::string selectedLabelUtf8;
	std::string noteUtf8;
};

// 解析并校验 request_user_input 的 UTF-8 JSON 参数。
bool ParseAIChatUserInputRequestArguments(
	const std::string& argumentsJsonUtf8,
	std::vector<AIChatUserInputQuestion>& outQuestions,
	std::string& outError);

// 校验 UI 返回的答案，并构建返回给模型的 UTF-8 JSON。
bool BuildAIChatUserInputResponseJson(
	const std::vector<AIChatUserInputQuestion>& questions,
	const std::vector<AIChatUserInputAnswer>& answers,
	std::string& outResponseJsonUtf8,
	std::string& outError);

// 构建 request_user_input 工具定义的 UTF-8 JSON，供各 AI 协议共用。
std::string BuildAIChatUserInputToolDefinitionJson();

// 构建带标准参数示例的 request_user_input 校验失败结果。
std::string BuildAIChatUserInputValidationErrorJson(const std::string& error);

// 构建 request_user_input 参数与结果协议的自检报告。
std::string BuildAIChatUserInputRequestSelfTestJson();
