#pragma once

// 易语言 IDE 编译输出内部捕获：安全解析跨版本输出函数，并提供按编译会话隔离的日志缓冲。

#include <cstdint>
#include <string>

namespace IdeCompileOutputCapture {

using SessionId = std::uint64_t;

struct CaptureSnapshot {
	std::string text;
	bool truncated = false;
};

// 在已开启的 Detours 事务中解析并附加 IDE 输出函数 Hook。
// 返回 false 时调用方应继续提交其他 Hook；编译工具会自动回退到控件取文本。
bool AttachToCurrentDetourTransaction();

// Detours 事务提交后同步安装状态。
void CompleteHookInstallation(bool transactionCommitted);

bool IsHookAvailable();

// 仅 Hook 可用时创建捕获会话；不可用时返回 0。
SessionId BeginCapture();
CaptureSnapshot SnapshotCapture(SessionId sessionId);
CaptureSnapshot EndCapture(SessionId sessionId);
void CancelCapture(SessionId sessionId);

// 无需启动 IDE 的解析器与会话缓冲自检。
std::string BuildSelfTestJson();

} // namespace IdeCompileOutputCapture
