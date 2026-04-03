#pragma once

#include <functional>
#include <string>

// PowerShell 命令执行结果。
struct PowerShellRunResult {
	bool ok = false;
	bool cancelled = false;
	bool timedOut = false;
	unsigned long exitCode = 0;
	std::string effectiveWorkingDirectory;
	std::string stdOut;
	std::string stdErr;
	std::string error;
};

// PowerShell 命令执行器。
class PowerShellToolRunner {
public:
	static PowerShellRunResult Run(
		const std::string& commandUtf8,
		const std::string& workingDirectoryUtf8,
		int timeoutSeconds,
		const std::function<bool()>& cancelCallback = {});
};
