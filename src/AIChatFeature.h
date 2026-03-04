#pragma once

#include <Windows.h>

class ConfigManager;

namespace AIChatFeature {

void Initialize(HWND mainWindow, ConfigManager* configManager);
void Shutdown();
void OpenDialog();
bool HandleMainWindowMessage(UINT uMsg, WPARAM wParam, LPARAM lParam);

} // namespace AIChatFeature

