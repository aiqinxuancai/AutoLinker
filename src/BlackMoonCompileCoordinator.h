#pragma once

#include "IDEFacade.h"
#include "..\\thirdparty\\json.hpp"
#include <filesystem>

// 黑月编译会话：持有跨进程配置锁，原子切换配置，并在结果收集后恢复。
class BlackMoonCompileCoordinator final {
public:
	BlackMoonCompileCoordinator() = default;
	~BlackMoonCompileCoordinator();
	BlackMoonCompileCoordinator(const BlackMoonCompileCoordinator&) = delete;
	BlackMoonCompileCoordinator& operator=(const BlackMoonCompileCoordinator&) = delete;
	bool Invoke(IDEFacade::CompileOutputKind kind, const std::string& outputPath,
		IDEFacade::BlackMoonCompileMode mode, std::string* normalizedPath, std::string* diagnostics);
	bool Finish(std::string& error);
	static nlohmann::json Inspect();
	// 校验并规范化工具参数；失败结果包含可供 AI 纠正请求的 hint。
	static bool ValidateArguments(nlohmann::json& args, nlohmann::json& error);
	static nlohmann::json SelfTest();

private:
	bool Begin(const std::filesystem::path& path, int mode, DWORD timeout, std::string& error);
	HANDLE mutex_ = nullptr;
	bool locked_ = false;
	bool changed_ = false;
	std::filesystem::path path_;
	std::string original_;
	std::string temporary_;
};
