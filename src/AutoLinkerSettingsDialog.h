#pragma once

// AutoLinker 统一设置窗口：提供左侧导航、页面深链和独立保存结果。

#include <Windows.h>

#include <string>

enum class AutoLinkerSettingsPageId {
	LastUsed = -1,
	AiService = 0,
	AiOther,
	Mcp,
	Skills,
	ChatTheme,
	ProjectAgents,
	Linker,
	EcSwitch,
	ForceLinkLib,
	ProjectBuild,
	LogOptimization,
	About,
	Count
};

struct AutoLinkerSettingsResult {
	bool aiSettingsSaved = false;
	bool mcpSettingsSaved = false;
	bool skillsChanged = false;
};

inline constexpr UINT WM_AUTOLINKER_SETTINGS_PAGE_SAVED = WM_APP + 0x3A1;

AutoLinkerSettingsResult ShowAutoLinkerSettingsDialog(
	HWND owner,
	AutoLinkerSettingsPageId initialPage = AutoLinkerSettingsPageId::LastUsed);

// 无需启动 IDE 的统一设置注册表、固定链接与开关依赖自检。
std::string BuildAutoLinkerSettingsSelfTestJson();
