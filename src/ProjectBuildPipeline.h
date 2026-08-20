#pragma once

// 项目配置异步编译流程：快照当前工程并交给独立 IDE 进程无头编译。

#include "ProjectBuildConfigManager.h"

#include <string>

struct ProjectBuildPipelineResult {
	bool ok = false;
	std::string stage;
	std::string message;
	std::string outputPath;
};

// 在调用线程完成当前工程快照，随后异步执行动作和无头编译。
ProjectBuildPipelineResult StartProjectBuildPipelineAsync(
	const std::filesystem::path& sourcePath,
	const ProjectBuildConfig& config);

