#pragma once
// EcSwitchConfigDialog.h - EC 模块动静态自动切换设置页

#include <Windows.h>

// 打开“EC 模块自动切换设置”WebView2 对话框。
void ShowEcSwitchConfigDialog(HWND owner);

// 创建统一设置窗口使用的 EC 模块切换子页。
HWND CreateEcSwitchConfigSettingsPage(HWND parent);
