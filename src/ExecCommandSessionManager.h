#pragma once

#include <cstddef>
#include <functional>
#include <string>

// 内部 AI 命令会话请求。
struct ExecCommandRequest {
	std::string ownerSessionId;
	std::string commandUtf8;
	std::string workingDirectoryUtf8;
	std::string shellUtf8;
	bool loginShell = true;
	bool tty = false;
	unsigned int yieldTimeMs = 10000;
	std::size_t maxOutputTokens = 10000;
};

// 内部 AI 命令会话轮询请求。
struct ExecCommandWriteRequest {
	std::string ownerSessionId;
	int sessionId = 0;
	std::string charsUtf8;
	unsigned int yieldTimeMs = 250;
	std::size_t maxOutputTokens = 10000;
};

// Codex 风格命令执行结果。
struct ExecCommandResult {
	bool ok = false;
	std::string responseUtf8;
	std::string error;
};

// 管理内部 AI 创建的后台命令会话及其增量输出。
class ExecCommandSessionManager {
public:
	static ExecCommandSessionManager& Instance();

	ExecCommandResult Execute(
		const ExecCommandRequest& request,
		const std::function<bool()>& cancelCallback = {});
	ExecCommandResult WriteStdin(
		const ExecCommandWriteRequest& request,
		const std::function<bool()>& cancelCallback = {});

	void TerminateOwnerSession(const std::string& ownerSessionId);
	void TerminateAll();

	// 构建无需易语言 IDE 的执行会话自检报告。
	std::string BuildSelfTestJson();

private:
	ExecCommandSessionManager();
	~ExecCommandSessionManager();
	ExecCommandSessionManager(const ExecCommandSessionManager&) = delete;
	ExecCommandSessionManager& operator=(const ExecCommandSessionManager&) = delete;

	class Impl;
	Impl* m_impl = nullptr;
};
