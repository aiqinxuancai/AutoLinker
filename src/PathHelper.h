#pragma once
#include <filesystem>
#include <optional>
#include <string>
#include <vector>
#include <Windows.h>

// 获取当前主程序（易语言IDE）所在目录路径。
std::string GetBasePath();

// 获取 AutoLinker 工作目录路径（主程序目录/AutoLinker）。
std::filesystem::path GetAutoLinkerDirectoryPath();

// 获取 AutoLinker 日志目录路径，若目录不存在则自动创建。
std::filesystem::path GetAutoLinkerLogDirectoryPath();

// 获取指定名称的日志文件完整路径；.log 文件会自动追加当前进程 PID，避免多实例互相覆盖。
std::filesystem::path GetAutoLinkerLogFilePath(const std::string& fileName);

// 清理日志目录中最后写入时间早于指定天数的 .log 文件。
void CleanupOldAutoLinkerLogFiles(int retentionDays = 3);

// 获取 AI 会话存储根目录路径，若不存在则自动创建。
std::filesystem::path GetAutoLinkerSessionRootDirectoryPath();

// 将文本清洗为合法的文件名/目录名片段，替换 Windows 非法字符为下划线。
std::string SanitizePathComponentForStorage(const std::string& text);

// 从文本中提取首对" - "分隔符之间的内容；找不到则返回空字符串。
std::string ExtractBetweenDashes(const std::string& text);

// 在文件中搜索指定字节序列，返回首次出现的字节偏移量；未找到则返回 std::nullopt。
std::optional<size_t> FindByteInFile(const std::string& filename, const std::vector<char>& search_bytes);

// 从链接器命令行中解析 /out:"..." 参数，返回输出文件名（不含目录）。
std::string GetLinkerCommandOutFileName(const std::string& s);

// 从链接器命令行中提取 krnln 静态库的完整路径。
std::string GetLinkerCommandKrnlnFileName(const std::string& s);
