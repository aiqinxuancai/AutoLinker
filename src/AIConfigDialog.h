#pragma once

#include <Windows.h>
#include <string>
#include <vector>

#include "AIService.h"
#include "AIJsonConfig.h"

enum class AIPreviewAction {
	Cancel = 0,
	PrimaryConfirm = 1,
	SecondaryConfirm = 2,
	TertiaryConfirm = 3
};

bool ShowAIConfigDialog(HWND owner, AIJsonConfig& jsonConfig, AISettings& ioSettings);

// 创建统一设置窗口使用的 AI 模型子页，返回值由父窗口负责销毁。
HWND CreateAIConfigSettingsPage(HWND parent);

// 打开“AutoLinker 链接器设置”对话框（WebView2），用于查看/编辑 AutoLinker/Config 下的 link.ini 配置。
void ShowLinkerConfigDialog(HWND owner);
HWND CreateLinkerConfigSettingsPage(HWND parent);
// 打开“AutoLinker AI 对话配色设置”对话框（WebView2）。
void ShowAIChatThemeConfigDialog(HWND owner);
HWND CreateAIChatThemeConfigSettingsPage(HWND parent);
AIPreviewAction ShowAIPreviewDialogEx(
	HWND owner,
	const std::string& title,
	const std::string& content,
	const std::string& primaryText = "",
	const std::string& secondaryText = "",
	const std::string& tertiaryText = "");
bool ShowAIPreviewDialog(HWND owner, const std::string& title, const std::string& content, const std::string& confirmText = "");
bool ShowAITextInputDialog(HWND owner, const std::string& title, const std::string& hint, std::string& ioText);
