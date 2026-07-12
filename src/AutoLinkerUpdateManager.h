#pragma once

// AutoLinker 支持库更新管理器：检查 Release，并在 IDE 退出后替换已加载的 fne。
namespace AutoLinkerUpdateManager {

// 在后台检查并执行 AutoLinker.fne 更新流程。
void RunUpdateInBackground();

} // namespace AutoLinkerUpdateManager
