#pragma once

#include <cstdint>
#include <filesystem>
#include <string>
#include <vector>
#include <memory>
#include <atomic>

// 会话级工程镜像：把当前 IDE 内存源码用 e-packager 解包到临时目录，并提供文件到真实程序项的映射。
namespace WorkspaceMirror {

enum class RefreshMode {
	Auto,
	MainOnly,
	Full
};

// 解包文件对应的真实程序项。
struct ProgramItemRef {
	std::string relativePathUtf8;
	std::string pageNameLocal;
	std::string kind;
	bool editable = false;
	bool fixedTable = false;
	bool formXml = false;
};

// 已在主线程准备完成的只读镜像快照；后续文件 IO 可安全地在工作线程执行。
struct FileAccessSnapshot {
	std::filesystem::path mirrorRoot;
	std::vector<std::string> relativePathsUtf8;
	std::uint64_t generation = 0;
	// 读取期间持有目录所有权，刷新不能删除仍在使用的旧镜像。
	std::shared_ptr<void> lifetime;
};

// 外部易语言源码解包结果；目录路径相对于当前工程镜像根目录。
struct ReferenceUnpackResult {
	std::string sourcePathUtf8;
	std::string outputDirectoryUtf8;
	std::uint64_t fileCount = 0;
	std::uint64_t generation = 0;
};

// 分阶段刷新任务：主线程采集/发布，后台只处理不可变输入和候选目录。
struct RefreshTask;
std::shared_ptr<RefreshTask> CaptureRefreshInputOnMainThread(RefreshMode mode, std::string& error);
void BuildRefreshCandidate(const std::shared_ptr<RefreshTask>& task);
bool CommitRefreshOnMainThread(const std::shared_ptr<RefreshTask>& task, std::string& error, std::string& mode);
void CancelRefresh(const std::shared_ptr<RefreshTask>& task);
// 参考源码和增量写入同样生成独立候选目录，读者持有的版本保持不变。
std::shared_ptr<RefreshTask> CaptureReferenceInputOnMainThread(const std::string& pathUtf8, std::string& error);
std::shared_ptr<RefreshTask> CaptureTextUpdateOnMainThread(const std::string& pathUtf8, const std::string& textLocal, std::string& error);
ReferenceUnpackResult GetReferenceResult(const std::shared_ptr<RefreshTask>& task);

// 源码关闭/切换的独立代次，不能以路径比较替代。
std::uint64_t GetContextEpoch();

// 访问当前源码的接口须在 IDE 主线程调用；源码变化通知在镜像加锁前完成。
// 校验当前源码的镜像已准备；不会隐式访问文件或启动解包。
bool EnsureMirrorFresh(std::string& outError);

// 校验文件访问环境；调用前由调度器完成必要的异步刷新。
bool PrepareFileAccess(std::string& outError);

// 获取已经准备好的只读镜像快照；此函数不会访问 IDE 或自动刷新。
bool GetPreparedFileAccessSnapshot(FileAccessSnapshot& outSnapshot, std::string& outError);

// 基于已准备快照解析相对路径；不会访问 IDE 或改变镜像状态。
bool ResolvePreparedFilePath(
	const FileAccessSnapshot& snapshot,
	const std::string& filePathUtf8,
	std::filesystem::path& outFullPath,
	std::string& outRelativePathUtf8,
	std::string& outError);

// 标记镜像已过期，下一次读取/搜索/列出时惰性重建。
void InvalidateMirror();

// 清理当前镜像目录并重置状态。
void ResetAndCleanup();

// 获取当前镜像根目录。
bool GetMirrorRoot(std::filesystem::path& outRoot, std::string& outError);

// 获取当前镜像代次；刷新、失效、增量写入或重置时都会变化。
std::uint64_t GetGeneration();

// 解析镜像内相对路径到真实文件路径。
bool ResolveFilePath(
	const std::string& filePathUtf8,
	std::filesystem::path& outFullPath,
	std::string& outRelativePathUtf8,
	std::string& outError);

// 解析镜像内源码文件到 IDE 真实程序项；只允许当前工程 src 下可写源码文件。
bool ResolveFileToProgramItem(
	const std::string& filePathUtf8,
	ProgramItemRef& outItem,
	std::string& outError);

// 返回当前镜像中可见的文件列表。
bool ListMirrorFiles(std::vector<std::string>& outRelativePathsUtf8, std::string& outError);

// 构建镜像目录所有者解析和相对路径校验自检报告。
std::string BuildSelfTestReportJson();

} // namespace WorkspaceMirror
