#include "AIChatSourceBackup.h"

#include <Windows.h>
#include <algorithm>
#include <atomic>
#include <format>
#include <vector>

namespace AIChatSourceBackup {
namespace {
struct TemporarySnapshot {
    std::filesystem::path directory;
    ~TemporarySnapshot() { std::error_code ec; std::filesystem::remove_all(directory, ec); }
};

bool IsBackupName(const std::wstring& name)
{
    const auto dot = name.rfind(L'.');
    if (dot == std::wstring::npos || (name.substr(dot) != L".e" && name.substr(dot) != L".ec")) return false;
    if (dot < 18 || name.substr(14, 2) != L"-[") return false;
    for (size_t i = 0; i < 14; ++i) if (name[i] < L'0' || name[i] > L'9') return false;
    const auto close = name.find(L']', 16);
    if (close == std::wstring::npos || close == 16 || close >= dot) return false;
    if (close + 1 == dot) return true;
    if (name[close + 1] != L'-' || close + 2 == dot) return false;
    for (size_t i = close + 2; i < dot; ++i) if (name[i] < L'0' || name[i] > L'9') return false;
    return true;
}

void Prune(const std::filesystem::path& directory, const std::filesystem::path& newest, int keepCount)
{
    struct Entry { std::filesystem::path path; std::filesystem::file_time_type time; };
    std::vector<Entry> entries;
    for (const auto& entry : std::filesystem::directory_iterator(directory)) {
        if (!entry.is_regular_file() || entry.is_symlink() || !IsBackupName(entry.path().filename().wstring())) continue;
        entries.push_back({entry.path(), entry.last_write_time()});
    }
    std::sort(entries.begin(), entries.end(), [&](const Entry& a, const Entry& b) {
        if (a.path == newest || b.path == newest) return a.path == newest && b.path != newest;
        if (a.time != b.time) return a.time > b.time;
        return a.path.filename() > b.path.filename();
    });
    for (size_t i = static_cast<size_t>(keepCount); i < entries.size(); ++i) {
        std::filesystem::remove(entries[i].path);
    }
}
} // namespace

std::wstring BuildRequestLabel(std::wstring_view request)
{
    std::wstring result;
    size_t offset = 0;
    for (int count = 0; offset < request.size() && count < 20; ++count) {
        const wchar_t ch = request[offset++];
        if (ch >= 0xD800 && ch <= 0xDBFF && offset < request.size() &&
            request[offset] >= 0xDC00 && request[offset] <= 0xDFFF) {
            result += ch; result += request[offset++];
        } else if (ch < 32 || (ch >= 0xD800 && ch <= 0xDFFF) ||
            std::wstring_view(L"<>:\"/\\|?*[]").find(ch) != std::wstring_view::npos) {
            result += L'_';
        } else {
            result += ch;
        }
    }
    if (offset < request.size()) result += L"...";
    if (result.empty()) result = L"图片消息";
    return result;
}

bool CreateBackup(const std::filesystem::path& source, std::wstring_view request,
    int keepCount, const SnapshotWriter& writer, std::filesystem::path& output, std::string& error, bool pruneAfterWrite)
{
    output.clear(); error.clear();
    try {
        if (!source.is_absolute() || !std::filesystem::is_regular_file(source)) {
            error = "source_file_not_saved_to_disk"; return false;
        }
        const auto extension = source.extension().wstring();
        if (_wcsicmp(extension.c_str(), L".e") != 0 && _wcsicmp(extension.c_str(), L".ec") != 0) {
            error = "unsupported_source_extension"; return false;
        }
        keepCount = std::clamp(keepCount, 1, kMaxKeepCount);
        const auto directory = source.parent_path() / (source.filename().wstring() + L".BAK");
        std::filesystem::create_directories(directory);
        // 临时路径不含用户需求中的 Unicode；最终用宽字符 API 发布备份。
        static std::atomic<unsigned long long> sequence{0};
        TemporarySnapshot temporary;
        for (;;) {
            auto candidate = std::filesystem::temp_directory_path() /
                std::format(L"AutoLinkerBackup-{}-{}-{}", GetCurrentProcessId(), GetTickCount64(), ++sequence);
            if (std::filesystem::create_directory(candidate)) { temporary.directory = candidate; break; }
        }
        const auto snapshot = temporary.directory / (L"snapshot" + extension);
        SYSTEMTIME time{}; GetLocalTime(&time);
        const auto stem = std::format(L"{:04}{:02}{:02}{:02}{:02}{:02}-[{}]",
            time.wYear, time.wMonth, time.wDay, time.wHour, time.wMinute, time.wSecond, BuildRequestLabel(request));
        if (!writer(snapshot, error)) return false;
        if (!std::filesystem::is_regular_file(snapshot) || std::filesystem::file_size(snapshot) == 0) {
            error = "empty_snapshot"; return false;
        }
        std::filesystem::last_write_time(snapshot, std::filesystem::file_time_type::clock::now());
        for (unsigned long long suffix = 0;; ++suffix) {
            auto target = directory / (stem + (suffix == 0 ? L"" : L"-" + std::to_wstring(suffix)) +
                (_wcsicmp(extension.c_str(), L".ec") == 0 ? L".ec" : L".e"));
            // 不允许覆盖：同秒、同需求提交也保留独立快照。
            if (MoveFileExW(snapshot.c_str(), target.c_str(), MOVEFILE_COPY_ALLOWED)) { output = target; break; }
            const DWORD code = GetLastError();
            if (code != ERROR_FILE_EXISTS && code != ERROR_ALREADY_EXISTS) {
                error = "publish_snapshot_failed: " + std::to_string(code); return false;
            }
        }
        if (pruneAfterWrite) Prune(directory, output, keepCount);
        return true;
    } catch (const std::exception& ex) {
        error = ex.what();
        // 清理失败时已发布的新快照仍然有效，调用方同时报告警告。
        return !output.empty();
    }
}
std::vector<std::filesystem::path> ListBackups(const std::filesystem::path& source)
{
    std::vector<std::filesystem::path> result;
    const auto directory = source.parent_path() / (source.filename().wstring() + L".BAK");
    if (!std::filesystem::exists(directory)) return result;
    for (const auto& entry : std::filesystem::directory_iterator(directory)) {
        if (entry.is_regular_file() && !entry.is_symlink() && IsBackupName(entry.path().filename().wstring()))
            result.push_back(entry.path());
    }
    std::sort(result.begin(), result.end(), [](const auto& a, const auto& b) {
        const auto ta = std::filesystem::last_write_time(a), tb = std::filesystem::last_write_time(b);
        return ta != tb ? ta > tb : a.filename() > b.filename();
    });
    return result;
}
void PruneBackups(const std::filesystem::path& source, const std::filesystem::path& newest, int keepCount)
{
    Prune(source.parent_path() / (source.filename().wstring() + L".BAK"), newest,
        std::clamp(keepCount, 1, kMaxKeepCount));
}
} // namespace AIChatSourceBackup
