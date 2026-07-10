#pragma once

#include <cstdint>
#include <filesystem>
#include <string>
#include <vector>

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
};

// 确保当前源码的解包镜像可用；必要时导出内存快照并重新解包。
bool EnsureMirrorFresh(std::string& outError);

// 在 IDE 主线程准备文件访问环境；只在镜像失效时触发快照和解包。
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

// 强制刷新当前源码镜像。Auto 保持原策略；MainOnly 只刷新源文件；Full 重建完整镜像。
bool RefreshMirror(std::string& outError, std::string* outMode = nullptr, RefreshMode mode = RefreshMode::Auto);

// 标记镜像已过期，下一次读取/搜索/列出时惰性重建。
void InvalidateMirror();

// 将一次成功写回 IDE 的源码同步到当前镜像文件，避免后续读搜重复解包。
bool UpdateMirrorTextFile(
	const std::string& filePathUtf8,
	const std::string& textLocal,
	std::string& outError);

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
