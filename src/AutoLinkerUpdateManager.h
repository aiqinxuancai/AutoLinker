#pragma once

#include <Windows.h>

#include "ComponentUpdateStatus.h"

// AutoLinker 支持库更新管理器：检查 Release，并在 IDE 退出后替换已加载的 fne。
namespace AutoLinkerUpdateManager {

// 仅在后台刷新最新版本状态，不下载或退出 IDE。
void CheckForUpdatesInBackground();
// 在后台检查并执行 AutoLinker.fne 更新流程。
void RunUpdateInBackground();
// 设置状态变化通知窗口；窗口销毁前应清空。
void SetStatusNotificationWindow(HWND window);
ComponentUpdateStatus GetStatus();

} // namespace AutoLinkerUpdateManager
