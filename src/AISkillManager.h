#pragma once

#include <filesystem>
#include <optional>
#include <string>
#include <vector>

#include "..\\thirdparty\\json.hpp"

// AI 技能作用域。
enum class AISkillScope {
	User,
	Repo
};

// 已发现的 AI 技能信息。
struct AISkillInfo {
	std::string name;
	std::string description;
	std::string shortDescription;
	AISkillScope scope = AISkillScope::User;
	std::filesystem::path directory;
	std::filesystem::path skillFile;
	bool enabled = true;
	bool valid = false;
	std::string error;
	std::string sourceRepository;
	std::string sourceRef;
	std::string sourcePath;
	std::string sourceCommit;
	long long installedAtUnixMs = 0;
};

// GitHub 仓库中可安装的技能候选项。
struct AISkillInstallCandidate {
	std::string name;
	std::string description;
	std::string repositoryPath;
};

// AI 技能发现、读取和安装管理。
namespace AISkillManager {

// 返回全局技能根目录。
std::filesystem::path GetGlobalSkillsRoot();
// 返回当前项目技能根目录；没有打开项目时返回空。
std::optional<std::filesystem::path> GetProjectSkillsRoot();
// 发现全局和当前项目技能。
std::vector<AISkillInfo> DiscoverSkills();
// 构建设置页初始数据。
std::string BuildSettingsPayloadJson();
// 构建本轮系统提示词中的技能目录和显式技能正文。
std::string BuildRuntimePromptAddon(const std::string& latestUserMessage);
// 执行模型侧的受限技能资源读取工具。
std::string ExecuteReadSkillResourceTool(const std::string& argumentsJson, bool& outOk);

// 搜索 skills.sh。
std::string SearchSkillsSh(const std::string& query);
// 读取 skills.sh 技能详情。
std::string GetSkillsShSkillDetails(const std::string& source, const std::string& skillId);
// 检查公开 GitHub 仓库并列出技能候选项。
std::string InspectGitHubRepository(const std::string& source);
// 从公开 GitHub 仓库安装一个技能。
std::string InstallFromGitHub(
	const std::string& source,
	AISkillScope scope,
	const std::string& selectedRepositoryPath,
	const std::string& expectedSkillId,
	bool allowReplace);
// 更新一个由 AutoLinker 安装的技能。
std::string UpdateInstalledSkill(const std::string& skillFilePath);
// 删除一个受管技能目录。
std::string RemoveInstalledSkill(const std::string& skillFilePath);
// 修改技能启用状态。
bool SetSkillEnabled(const std::string& skillFilePath, bool enabled, std::string& outError);

// 无需启动 IDE 的技能核心自检。
std::string BuildSelfTestJson();

} // namespace AISkillManager
