#pragma once

#include <Windows.h>

#include <condition_variable>
#include <mutex>
#include <string>

class ConfigManager;

// 主线程工具执行请求。
struct ToolExecutionRequest {
	std::string toolName;
	std::string argumentsJson;
	std::string resultJson;
	bool ok = false;
	bool done = false;
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
	bool accepted = false;          // 用户点击了主确认按钮。
	bool secondaryAccepted = false; // 用户点击了次确认按钮。
	bool done = false;
	std::mutex mutex;
	std::condition_variable cv;
};

// 获取 AI 对话主窗口句柄。
HWND GetAIChatMainWindowForTooling();
// 获取 AI 对话配置管理器。
ConfigManager* GetAIChatConfigManagerForTooling();
// 获取 AI 对话工具执行消息。
UINT GetAIChatToolExecMessageForTooling();
// 请求确认对话。outSecondaryAccepted 为次确认按钮点击结果。
bool RequestConfirmationForTooling(
	const std::string& title,
	const std::string& content,
	const std::string& primaryText,
	const std::string& secondaryText,
	bool& outAccepted,
	bool& outSecondaryAccepted);
