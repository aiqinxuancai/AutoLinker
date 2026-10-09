#pragma once

#include <filesystem>
#include <functional>
#include <string>
#include <string_view>
#include <vector>

class AIJsonConfig;

// 用户发送消息时的源码快照备份；文件名和文件操作始终使用 Unicode。
namespace AIChatSourceBackup {
// 三类快照独立保留；旧文件名归入对话类。
enum class Kind { Chat, Timed, ToolWrite };
inline constexpr char kTimedEnabledKey[] = "timed_source_backup_enabled";
inline constexpr char kTimedKeepKey[] = "timed_source_backup_keep_count";
inline constexpr char kIntervalKey[] = "timed_source_backup_interval_minutes";
inline constexpr char kToolEnabledKey[] = "tool_source_backup_enabled";
inline constexpr char kToolKeepKey[] = "tool_source_backup_keep_count";
inline constexpr unsigned int kTimerId = 0xA1B;
bool IsEnabled(const AIJsonConfig& config, Kind kind);
int GetKeepCount(const AIJsonConfig& config, Kind kind);
int GetIntervalMinutes(const AIJsonConfig& config);
Kind GetBackupKind(const std::filesystem::path& backup);
// 主线程定时器轮询；忙碌时延后，不在工作线程调用 IDE。
void Tick(const AIJsonConfig* config, bool busy);
void CaptureBeforeTool(const AIJsonConfig* config, const std::string& toolName);
inline constexpr char kEnabledKey[] = "ai_source_backup_enabled";
inline constexpr char kKeepCountKey[] = "ai_source_backup_keep_count";
inline constexpr int kDefaultKeepCount = 10;
inline constexpr int kMaxKeepCount = 1000;

// 默认启用；无效保留数量回退到 10。
bool IsEnabled(const AIJsonConfig& config);
int GetKeepCount(const AIJsonConfig& config);

// 在 IDE 主线程抓取发送时的工程状态；失败仅记录日志，不阻止消息发送。
void CaptureBeforeSend(const AIJsonConfig* config, const std::string& requestUtf8);

// 按 Unicode 码点取前二十个字符，替换 Windows 非法字符，超长追加省略号。
std::wstring BuildRequestLabel(std::wstring_view request);

// 写入成功后才清理旧备份；回调写出完整工程文件，便于无 IDE 测试。
using SnapshotWriter = std::function<bool(const std::filesystem::path&, std::string&)>;
bool CreateBackup(const std::filesystem::path& source, std::wstring_view request,
    int keepCount, const SnapshotWriter& writer, std::filesystem::path& output, std::string& error,
    bool pruneAfterWrite = true, Kind kind = Kind::Chat);

// 列举本工程的自动备份，最新在前；恢复安全备份允许延后清理。
std::vector<std::filesystem::path> ListBackups(const std::filesystem::path& source);
void PruneBackups(const std::filesystem::path& source, const std::filesystem::path& newest, int keepCount);
} // namespace AIChatSourceBackup
