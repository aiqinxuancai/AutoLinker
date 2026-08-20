#pragma once

// 项目配置模型、文件读写和路径变量展开接口。

#include <filesystem>
#include <string>
#include <vector>

enum class ProjectBuildTarget {
	Auto,
	WinExe,
	WinConsoleExe,
	WinDll,
	Ecom
};

struct ProjectBuildConfig {
	std::string name;
	ProjectBuildTarget target = ProjectBuildTarget::Auto;
	bool staticCompile = false;
	std::string outputPath;
	std::vector<std::string> preBuildCommands;
	std::vector<std::string> postBuildCommands;
};

struct ProjectBuildConfigFile {
	int version = 1;
	std::string activeConfiguration;
	std::vector<ProjectBuildConfig> configurations;
};

struct ProjectBuildVariableContext {
	std::filesystem::path programDir;
	std::filesystem::path projectDir;
	std::string projectName;
	std::filesystem::path sourcePath;
	std::filesystem::path outputPath;
};

std::filesystem::path GetProjectBuildConfigPath(const std::filesystem::path& sourcePath);
ProjectBuildConfigFile LoadProjectBuildConfigFile(
	const std::filesystem::path& sourcePath,
	std::string* outError = nullptr);
bool SaveProjectBuildConfigFile(
	const std::filesystem::path& sourcePath,
	const ProjectBuildConfigFile& file,
	std::string* outError = nullptr);

const char* ProjectBuildTargetToString(ProjectBuildTarget target);
ProjectBuildTarget ProjectBuildTargetFromString(const std::string& value);

std::string ExpandProjectBuildVariables(
	const std::string& text,
	const ProjectBuildVariableContext& context,
	std::vector<std::string>* outUnknownVariables = nullptr);

// 构建项目配置读写及变量展开的无 IDE 自检报告。
std::string BuildProjectBuildConfigSelfTestJson();

