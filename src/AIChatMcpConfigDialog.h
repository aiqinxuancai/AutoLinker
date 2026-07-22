#pragma once

#include <Windows.h>

// 显示外部 MCP 服务器配置窗口。
bool ShowAIChatMcpConfigDialog(HWND owner);

// 创建统一设置窗口使用的 MCP 子页，返回值由父窗口负责销毁。
HWND CreateAIChatMcpConfigSettingsPage(HWND parent);
