#pragma once

#include <Windows.h>
#include <string>
#include <vector>

#include "AIService.h"
#include "AIJsonConfig.h"

enum class AIPreviewAction {
	Cancel = 0,
	PrimaryConfirm = 1,
	SecondaryConfirm = 2
};

bool ShowAIConfigDialog(HWND owner, AIJsonConfig& jsonConfig, AISettings& ioSettings);

// 打开“AutoLinker 链接器设置”对话框（WebView2），用于查看/编辑 AutoLinker/Config 下的 link.ini 配置。
void ShowLinkerConfigDialog(HWND owner);
AIPreviewAction ShowAIPreviewDialogEx(
	HWND owner,
	const std::string& title,
	const std::string& content,
	const std::string& primaryText = "",
	const std::string& secondaryText = "");
bool ShowAIPreviewDialog(HWND owner, const std::string& title, const std::string& content, const std::string& confirmText = "");
bool ShowAITextInputDialog(HWND owner, const std::string& title, const std::string& hint, std::string& ioText);
