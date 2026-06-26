#pragma once
// ForceLinkLibManager.h - 核心库函数重写强制链接规则配置

#include <filesystem>
#include <string>
#include <vector>

// 单条强制链接规则；字符串使用当前系统本地编码，便于直接参与 CreateProcessA 命令行处理。
struct ForceLinkLibRule {
	bool enabled = true;          // 是否启用该规则。
	std::string linkerName;       // 可选：仅当当前链接器名称包含此文本时生效。
	std::string libPath;          // 要追加到 krnln_static.lib 前面的 .lib 文件路径。
};

// 管理 AutoLinker/ForceLinkLib.ini。
class ForceLinkLibManager {
public:
	ForceLinkLibManager();

	// 返回全部规则。
	std::vector<ForceLinkLibRule> getRules() const;

	// 替换全部规则并保存到 ForceLinkLib.ini。
	bool replaceRules(const std::vector<ForceLinkLibRule>& rules, std::string* errorMessage = nullptr);

	// 返回配置文件路径。
	const std::filesystem::path& getConfigFilePath() const;

private:
	void loadConfig();
	bool saveConfig(std::string* errorMessage = nullptr) const;

	std::filesystem::path configFilePath;
	std::vector<ForceLinkLibRule> rules_;
};
