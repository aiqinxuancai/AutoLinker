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
void SetUpdateAvailable(const std::string& latestVersion);
void OnCurrentSourceFilePathChanged(const std::string& previousPath, const std::string& currentPath);
bool HandleMainWindowMessage(UINT uMsg, WPARAM wParam, LPARAM lParam);
bool ExecutePublicTool(const std::string& toolName, const std::string& argumentsJson, std::string& outResultJsonUtf8, bool& outOk);
// 更新 AI 对话中的任务计划卡片。
bool UpdatePlanFromTool(const std::string& argumentsJsonUtf8, std::string& outResultJsonLocal, bool& outOk);
// 重新读取并应用当前 AI 对话配色。
void ReloadTheme();

} // namespace AIChatFeature
