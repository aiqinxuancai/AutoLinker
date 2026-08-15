#pragma once

#include <Windows.h>

#include <string>

// 编译并发协调器：防止同一 IDE 重入编译，并串行化相同输出路径的跨进程编译。
namespace CompileConcurrencyManager {

class ProcessExecutionGuard final {
public:
	ProcessExecutionGuard();
	~ProcessExecutionGuard();

	ProcessExecutionGuard(const ProcessExecutionGuard&) = delete;
	ProcessExecutionGuard& operator=(const ProcessExecutionGuard&) = delete;

	bool Acquired() const;

private:
	bool acquired_ = false;
};

class OutputPathLock final {
public:
	OutputPathLock() = default;
	~OutputPathLock();

	OutputPathLock(const OutputPathLock&) = delete;
	OutputPathLock& operator=(const OutputPathLock&) = delete;

	bool Acquire(
		const std::string& normalizedOutputPath,
		DWORD timeoutMilliseconds,
		std::string& outError);
	void Release();
	DWORD WaitElapsedMilliseconds() const;

private:
	HANDLE mutex_ = nullptr;
	bool acquired_ = false;
	DWORD waitElapsedMilliseconds_ = 0;
};

// 构建进程内重入和跨进程输出锁的无 IDE 自检报告。
std::string BuildSelfTestJson();

} // namespace CompileConcurrencyManager
