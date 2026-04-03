#pragma once

#include <functional>
#include <string>

// 在主线程执行工具调用。
std::string ExecuteToolCallOnMainThread(const std::string& toolName, const std::string& argumentsJson, bool& outOk);

// 执行一个 AI 工具调用，可选输出日志并支持取消。
std::string ExecuteToolCall(
	const std::string& toolName,
	const std::string& argumentsJson,
	bool& outOk,
	bool enableLog = true,
	const std::function<bool()>& cancelCallback = {});
