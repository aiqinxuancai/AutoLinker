#pragma once

#include <Windows.h>

#include <condition_variable>
#include <memory>
#include <mutex>
#include <string>

class AIJsonConfig;
class ConfigManager;

// 主线程工具执行请求。
struct ToolExecutionRequest {
	std::string toolName;
	std::string argumentsJson;
	std::string resultJson;
	bool bypassInteractiveApproval = false; // 19207 外部 MCP 调用不弹出交互审批。
	bool ok = false;
	bool done = false;
	bool cancelled = false;
	std::mutex mutex;
	std::condition_variable cv;
};

// 主线程工具对话请求。
struct ToolDialogRequest {
	enum class Kind {
		Confirmation
	};

	Kind kind = Kind::Confirmation;
	std::string title;
	std::string content;
	std::string primaryText;
	std::string secondaryText;
	std::string tertiaryText;
	bool accepted = false;          // 用户点击了主确认按钮。
	bool secondaryAccepted = false; // 用户点击了次确认按钮。
	bool tertiaryAccepted = false;  // 用户点击了第三确认按钮。
	bool done = false;
	std::mutex mutex;
	std::condition_variable cv;
};

// 获取 AI 对话主窗口句柄。
HWND GetAIChatMainWindowForTooling();
// 获取 AI 对话配置管理器。
ConfigManager* GetAIChatConfigManagerForTooling();
// 获取 AI JSON 配置。
AIJsonConfig* GetAIChatAIJsonConfigForTooling();
// 获取 AI 对话工具执行消息。
UINT GetAIChatToolExecMessageForTooling();
// 注册主线程工具请求，返回用于窗口消息交接的请求编号。
UINT_PTR RegisterToolExecutionRequestForTooling(const std::shared_ptr<ToolExecutionRequest>& request);
// 取消尚未被主线程接管的工具请求。
void CancelToolExecutionRequestForTooling(UINT_PTR requestId);
// 请求确认对话。outSecondaryAccepted 为次确认按钮点击结果。
bool RequestConfirmationForTooling(
	const std::string& title,
	const std::string& content,
	const std::string& primaryText,
	const std::string& secondaryText,
	bool& outAccepted,
	bool& outSecondaryAccepted,
	const std::string& tertiaryText = "",
	bool* outTertiaryAccepted = nullptr);
