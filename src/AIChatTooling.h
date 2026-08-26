#pragma once

#include <functional>
#include <string>

class HttpRequestCancellation;

// 在主线程执行工具调用；enableMcpAutoSave 仅由外部 MCP 请求开启。
std::string ExecuteToolCallOnMainThread(
	const std::string& toolName,
	const std::string& argumentsJson,
	bool& outOk,
	bool enableMcpAutoSave = false);

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

// 记录绕过常规执行器的内部 AI 工具请求，沿用统一的脱敏与摘要策略。
void LogAIChatToolRequest(const std::string& toolName, const std::string& argumentsJson);

// 记录绕过常规执行器的内部 AI 工具结果，结果文本使用本地编码。
void LogAIChatToolResponse(
	const std::string& toolName,
	const std::string& resultJsonLocal,
	double elapsedMs);

// 终止并清理指定内部聊天会话创建的全部命令进程。
void CloseInternalExecSession(const std::string& sessionId);

// 终止插件内全部内部 AI 命令进程。
void ShutdownInternalExecSessions();

// 判断调用域是否来自无需交互审批的 19207 外部 MCP 会话。
bool ShouldBypassToolApprovalForScope(const std::string& approvalScope);
