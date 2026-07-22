#pragma once

// 组件更新状态：供后台更新管理器与统一设置“关于”页交换线程安全快照。

#include <Windows.h>

#include <string>

enum class ComponentUpdateState {
	Idle,
	Checking,
	UpToDate,
	UpdateAvailable,
	Downloading,
	Installing,
	ReadyToRestart,
	Completed,
	Error
};

struct ComponentUpdateStatus {
	ComponentUpdateState state = ComponentUpdateState::Idle;
	std::string currentVersion;
	std::string latestVersion;
	std::string message;
	int progressPercent = -1;
};

inline constexpr UINT WM_AUTOLINKER_COMPONENT_UPDATE_STATUS = WM_APP + 0x3A2;

