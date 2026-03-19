#pragma once

#include <Windows.h>

#include <condition_variable>
#include <mutex>
#include <string>

struct ToolExecutionRequest {
	std::string toolName;
	std::string argumentsJson;
	std::string resultJson;
	bool ok = false;
	bool done = false;
	std::mutex mutex;
	std::condition_variable cv;
};

HWND GetAIChatMainWindowForTooling();
UINT GetAIChatToolExecMessageForTooling();
bool RequestCodeEditForTooling(
	const std::string& title,
	const std::string& hint,
	const std::string& initialCode,
	std::string& outCode);
