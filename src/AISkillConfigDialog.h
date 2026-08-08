#pragma once

#include <Windows.h>

// 创建统一设置窗口使用的 SKILL 管理子页，返回值由父窗口负责销毁。
HWND CreateAISkillConfigSettingsPage(HWND parent);

