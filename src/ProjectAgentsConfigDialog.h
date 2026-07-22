#pragma once
// ProjectAgentsConfigDialog.h - 当前易程序 AGENTS.md 项目规范设置页

#include <Windows.h>

// 打开“当前程序 AGENTS.md 设置”WebView2 对话框。
void ShowProjectAgentsConfigDialog(HWND owner);

// 创建统一设置窗口使用的当前项目 AGENTS.md 子页。
HWND CreateProjectAgentsConfigSettingsPage(HWND parent);

