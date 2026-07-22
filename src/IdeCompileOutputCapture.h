#pragma once

// 易语言 IDE 编译输出内部捕获：安全解析跨版本输出函数，并提供按编译会话隔离的日志缓冲。

#include <Windows.h>

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

// 控制调试文本异步合并输出；关闭时完全旁路快速输出路径。
void SetDebugOutputOptimizationEnabled(bool enabled) noexcept;
bool IsDebugOutputOptimizationEnabled() noexcept;

// 仅 Hook 可用时创建捕获会话；不可用时返回 0。
SessionId BeginCapture();
CaptureSnapshot SnapshotCapture(SessionId sessionId);
CaptureSnapshot EndCapture(SessionId sessionId);
void CancelCapture(SessionId sessionId);

// 在 IDE 主线程中处理调试输出合并消息和定时刷新。
bool HandleMainWindowMessage(HWND window, UINT message, WPARAM wParam, LPARAM lParam) noexcept;

// IDE 主窗口销毁时停止尚未执行的调试输出刷新。
void Shutdown(HWND window) noexcept;

// 无需启动 IDE 的解析器与会话缓冲自检。
std::string BuildSelfTestJson();

} // namespace IdeCompileOutputCapture
