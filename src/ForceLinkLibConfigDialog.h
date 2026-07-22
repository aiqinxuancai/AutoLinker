#pragma once
// ForceLinkLibConfigDialog.h - 核心库函数重写强制链接设置页

#include <Windows.h>

// 打开“核心库函数重写设置”WebView2 对话框。
void ShowForceLinkLibConfigDialog(HWND owner);

// 创建统一设置窗口使用的核心库函数重写子页。
HWND CreateForceLinkLibConfigSettingsPage(HWND parent);
