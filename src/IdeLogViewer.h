#pragma once

// 覆盖 IDE 原生日志控件并基于 WebView2 展示日志的高性能查看器。

#include <Windows.h>

#include <string>

class ConfigManager;

namespace IdeLogViewer {

// 设置日志中心开关状态使用的持久化配置。
void ConfigurePersistence(ConfigManager* configManager);
// 按持久化状态恢复日志中心；未启用时不创建界面。
void RestorePersistedState(HWND mainWindow);
bool Initialize(HWND mainWindow);
// 暂停并隐藏日志中心，露出下方的 IDE 原生日志控件。
void Close();
// IDE 退出时彻底释放查看器。
void Shutdown();
bool IsOpen();

// 验证嵌入式前端资源包含虚拟列表、普通搜索和正则搜索能力。
std::string BuildSelfTestJson();

} // namespace IdeLogViewer
