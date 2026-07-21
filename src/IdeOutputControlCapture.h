#pragma once

// IDE 日志控件采集器：通过窗口子类化捕获支持库写入输出控件的文本消息。

#include <Windows.h>

#include <string>

namespace IdeOutputControlCapture {

// 子类化指定的 IDE 日志控件；布局变化时向观察窗口投递指定消息。
bool Attach(HWND outputWindow, HWND observerWindow = nullptr, UINT layoutChangedMessage = 0) noexcept;

// 解除当前日志控件子类化。
void Detach() noexcept;

// 判断当前是否已附加到指定日志控件。
bool IsAttachedTo(HWND outputWindow) noexcept;

// 验证发送缓冲在原窗口过程处理后被复用时，已采集文本仍保持独立。
std::string BuildSelfTestJson();

} // namespace IdeOutputControlCapture
