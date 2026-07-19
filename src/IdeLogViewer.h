#pragma once

// 基于 FN_ADD_TAB 与 WebView2 的 IDE 高性能日志查看器。

#include <Windows.h>

#include <string>

namespace IdeLogViewer {

bool Initialize(HWND mainWindow);
// 暂停并隐藏日志中心，但保留首次 FN_ADD_TAB 建立的 IDE 页面映射。
void Close();
// IDE 退出时彻底释放查看器。
void Shutdown();
bool IsOpen();

// 验证嵌入式前端资源包含虚拟列表、普通搜索和正则搜索能力。
std::string BuildSelfTestJson();

} // namespace IdeLogViewer
