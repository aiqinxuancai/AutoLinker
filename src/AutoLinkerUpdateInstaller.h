#pragma once

#include <Windows.h>

#include <filesystem>
#include <string>

// 退出后更新器所需的已暂存文件和目标信息。
struct AutoLinkerUpdateInstallRequest {
	unsigned long processId = 0;
	std::filesystem::path stagedFne;
	std::filesystem::path targetFne;
	std::filesystem::path stagingRoot;
	std::filesystem::path logPath;
	std::string targetVersion;
};

// 退出后安装器：等待 IDE 进程结束，再原子替换 AutoLinker.fne。
class AutoLinkerUpdateInstaller {
public:
	static bool Launch(
		const AutoLinkerUpdateInstallRequest& request,
		PROCESS_INFORMATION& outProcess,
		std::string& outError);
};
