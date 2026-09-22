#pragma once

#include <string_view>

// 工具执行及本机转发的统一等待预算，单位为毫秒。
namespace ToolExecutionTimeouts {
inline constexpr int kConnectMs = 2000;
inline constexpr int kSendMs = 5000;
inline constexpr int kDefaultResponseMs = 30000;
inline constexpr int kCompileMs = 600000;
inline constexpr int kCompileResponseMs = kCompileMs + 30000;

constexpr int ResponseMs(std::string_view toolName)
{
	return toolName == "compile_with_output_path" ? kCompileResponseMs : kDefaultResponseMs;
}
}
