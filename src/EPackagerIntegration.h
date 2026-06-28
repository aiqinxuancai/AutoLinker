#pragma once

#include <filesystem>
#include <string>
#include <vector>

// e-packager 集成：负责工具更新、解包当前易语言工程和打开输出目录。
namespace EPackagerIntegration {

// e-packager 子进程执行结果。
struct ProcessRunResult {
	bool ok = false;
	unsigned long exitCode = 0;
	std::string stdOutBytes;
	std::string stdErrBytes;
	std::string error;
};

// 弹出目录选择窗口，并把当前 .e 文件解包到所选目录下的 {文件名}.unpack。
void RunCurrentSourceUnpackToDirectory();

// 构建“将[xxx.e]反编译到目录”菜单标题。
std::wstring BuildUnpackMenuTitle();

// 当前是否有可解包的 .e 文件。
bool CanUnpackCurrentSource();

// 获取当前打开的 .e 源文件路径。
bool GetCurrentSourcePath(std::filesystem::path& outSourcePath, std::string& outError);

// 构建当前工程内存快照默认路径。
std::filesystem::path BuildCurrentProjectSnapshotPathForSource(const std::filesystem::path& sourcePath);

// 写出当前 IDE 内存中的工程快照，包含未保存修改。
bool WriteCurrentProjectSnapshot(
	const std::filesystem::path& snapshotPath,
	size_t& outBytesWritten,
	std::string& outTrace,
	std::string& outError);

// 安全清理快照临时目录。
void CleanupSnapshotRoot(const std::filesystem::path& snapshotRoot);

// 确认 e-packager.exe 可用，必要时下载/更新。
bool EnsureToolReady(std::filesystem::path& outToolPath, std::string& outError);

// 强制检查最新版 e-packager.exe，失败时不静默回退到旧版本。
bool EnsureLatestToolReady(std::filesystem::path& outToolPath, std::string& outError);

// 后台检查并更新 e-packager.exe，忽略自动检查间隔。
void RunToolUpdateInBackground();

// 执行 e-packager 子进程并捕获输出。
ProcessRunResult RunProcessAndCapture(
	const std::filesystem::path& exePath,
	const std::vector<std::wstring>& args,
	const std::filesystem::path& workingDirectory);

} // namespace EPackagerIntegration
