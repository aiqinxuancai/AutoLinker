#pragma once

#include <filesystem>
#include <functional>
#include <string>
#include <string_view>

class AIJsonConfig;

// 用户发送消息时的源码快照备份；文件名和文件操作始终使用 Unicode。
namespace AIChatSourceBackup {
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
    int keepCount, const SnapshotWriter& writer, std::filesystem::path& output, std::string& error);
} // namespace AIChatSourceBackup
