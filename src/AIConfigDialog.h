#pragma once

#include <Windows.h>
#include <string>

#include "AIService.h"

bool ShowAIConfigDialog(HWND owner, AISettings& ioSettings);
bool ShowAIPreviewDialog(HWND owner, const std::string& title, const std::string& content, const std::string& confirmText = "");
bool ShowAITextInputDialog(HWND owner, const std::string& title, const std::string& hint, std::string& ioText);
bool ShowAICodeEditDialog(HWND owner, const std::string& title, const std::string& initialCode, std::string& ioCode);
