#pragma once

#include <Windows.h>
#include <string>

class AIJsonConfig;
class ConfigManager;

namespace AIChatFeature {

void Initialize(HWND mainWindow, ConfigManager* configManager, AIJsonConfig* aiJsonConfig);
void Shutdown();
void EnsureTabCreated();
void ActivateTab();
void OpenDialog();
bool HandleMainWindowMessage(UINT uMsg, WPARAM wParam, LPARAM lParam);
bool ExecutePublicTool(const std::string& toolName, const std::string& argumentsJson, std::string& outResultJsonUtf8, bool& outOk);

} // namespace AIChatFeature
