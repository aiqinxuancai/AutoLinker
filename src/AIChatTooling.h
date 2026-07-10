#pragma once

#include <functional>
#include <string>

class HttpRequestCancellation;

// 在主线程执行工具调用。
std::string ExecuteToolCallOnMainThread(const std::string& toolName, const std::string& argumentsJson, bool& outOk);

// 构建编译产物指纹判定的无 IDE 自检报告。
std::string BuildCompileArtifactFingerprintSelfTestJson();

// 执行一个 AI 工具调用，可选输出日志并支持取消。
std::string ExecuteToolCall(
	const std::string& toolName,
	const std::string& argumentsJson,
	bool& outOk,
	bool enableLog = true,
	const std::function<bool()>& cancelCallback = {},
	HttpRequestCancellation* cancellation = nullptr);

// 在指定授权域内执行工具，避免高风险工具的“一次允许”跨调用来源泄漏。
std::string ExecuteToolCall(
	const std::string& toolName,
	const std::string& argumentsJson,
	bool& outOk,
	bool enableLog,
	const std::function<bool()>& cancelCallback,
	HttpRequestCancellation* cancellation,
	const std::string& approvalScope);

// 清除指定调用域中缓存的高风险工具授权。
void ClearToolApprovalScope(const std::string& approvalScope);
