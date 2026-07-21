#pragma once

// IDE 日志控件采集器：通过窗口子类化捕获支持库写入输出控件的文本消息。

#include <Windows.h>

namespace IdeOutputControlCapture {

// 子类化指定的 IDE 日志控件；布局变化时向观察窗口投递指定消息。
bool Attach(HWND outputWindow, HWND observerWindow = nullptr, UINT layoutChangedMessage = 0) noexcept;

// 解除当前日志控件子类化。
void Detach() noexcept;

// 判断当前是否已附加到指定日志控件。
bool IsAttachedTo(HWND outputWindow) noexcept;

} // namespace IdeOutputControlCapture
