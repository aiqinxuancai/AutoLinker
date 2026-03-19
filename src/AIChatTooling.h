#pragma once

#include <string>

std::string ExecuteToolCallOnMainThread(const std::string& toolName, const std::string& argumentsJson, bool& outOk);
std::string ExecuteToolCall(const std::string& toolName, const std::string& argumentsJson, bool& outOk);
