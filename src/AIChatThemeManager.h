#pragma once

#include <string>
#include <vector>

#include "..\\thirdparty\\json.hpp"

// AI 对话配色管理：负责默认配色、用户配色文件和当前选中配色的读写。
namespace AIChatThemeManager {

struct ThemeEntry {
	std::string id;
	std::string name;
	nlohmann::json colors;
	bool isDefault = false;
};

// 返回内置默认配色变量表。
nlohmann::json GetDefaultColors();

// 返回当前选中的配色，读取失败时自动回退默认配色。
ThemeEntry LoadCurrentTheme();

// 构建 WebView 设置页初始数据。
std::string BuildConfigPayloadJson();

// 保存 WebView 提交的配色列表与当前选中项。
bool SaveConfigPayload(const nlohmann::json& data, std::string& outMessage);

// 生成应用当前配色到 AI 对话 WebView 的脚本。
std::wstring BuildApplyCurrentThemeScript();

} // namespace AIChatThemeManager
