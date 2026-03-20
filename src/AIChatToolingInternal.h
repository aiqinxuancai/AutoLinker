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
		CodeEdit,
		Confirmation
	};

	Kind kind = Kind::CodeEdit;
	std::string title;
	std::string hint;
	std::string content;
	std::string primaryText;
	std::string secondaryText;
	std::string resultText;
	bool accepted = false;
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
// 请求代码编辑对话。
bool RequestCodeEditForTooling(
	const std::string& title,
	const std::string& hint,
	const std::string& initialCode,
	std::string& outCode);
// 请求确认对话。
bool RequestConfirmationForTooling(
	const std::string& title,
	const std::string& content,
	const std::string& primaryText,
	const std::string& secondaryText,
	bool& outAccepted);
