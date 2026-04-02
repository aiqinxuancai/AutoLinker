#pragma once

#include <Windows.h>
#include <string>

#include "AIService.h"

enum class AIPreviewAction {
	Cancel = 0,
	PrimaryConfirm = 1,
	SecondaryConfirm = 2
};

bool ShowAIConfigDialog(HWND owner, AISettings& ioSettings);
AIPreviewAction ShowAIPreviewDialogEx(
	HWND owner,
	const std::string& title,
	const std::string& content,
	const std::string& primaryText = "",
	const std::string& secondaryText = "");
bool ShowAIPreviewDialog(HWND owner, const std::string& title, const std::string& content, const std::string& confirmText = "");
bool ShowAITextInputDialog(HWND owner, const std::string& title, const std::string& hint, std::string& ioText);
