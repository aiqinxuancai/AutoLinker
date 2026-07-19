#pragma once

// 易语言 IDE 输出页签控制器：兼容 FN_ADD_TAB 的多种标准 Tab 宿主布局。

#include <Windows.h>

#include <string>

namespace IdeOutputTabController {

struct HiddenTabState {
	HWND tabWindow = nullptr;
	int itemIndex = -1;
	bool hidden = false;
};

// 暂时隐藏已由 FN_ADD_TAB 注册的可见页签，但保留 IDE 内部页面映射。
bool HideTabByCaption(
	HWND mainWindow,
	HWND pageWindow,
	const std::string& caption,
	HiddenTabState& state);

// 原位恢复此前隐藏的页签，不重复调用 FN_ADD_TAB。
bool RestoreHiddenTab(
	HWND pageWindow,
	const std::string& caption,
	HiddenTabState& state);

} // namespace IdeOutputTabController
