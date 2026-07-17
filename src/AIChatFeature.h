#pragma once

#include <Windows.h>
#include <functional>
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
bool ExecutePublicTool(
	const std::string& toolName,
	const std::string& argumentsJson,
	std::string& outResultJsonUtf8,
	bool& outOk);
// 执行公开工具，并允许调用方在关闭或断连时请求取消。
bool ExecutePublicTool(
	const std::string& toolName,
	const std::string& argumentsJson,
	std::string& outResultJsonUtf8,
	bool& outOk,
	const std::function<bool()>& cancelCallback);
// 执行公开工具，并将高风险工具授权限制在指定调用域。
bool ExecutePublicTool(
	const std::string& toolName,
	const std::string& argumentsJson,
	std::string& outResultJsonUtf8,
	bool& outOk,
	const std::function<bool()>& cancelCallback,
	const std::string& approvalScope);
// 更新 AI 对话中的任务计划卡片。
bool UpdatePlanFromTool(const std::string& argumentsJsonUtf8, std::string& outResultJsonLocal, bool& outOk);
// 构建计划模式审批兼容性自检结果。
std::string BuildPlanModeSelfTestJson();
// 构建活动会话检查点自动恢复自检结果。
std::string BuildCheckpointResumeSelfTestJson();
// 重新读取并应用当前 AI 对话配色。
void ReloadTheme();

} // namespace AIChatFeature
