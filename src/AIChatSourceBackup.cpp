#include "AIChatSourceBackup.h"

#include <Windows.h>
#include <charconv>
#include "AIJsonConfig.h"
#include "EPackagerIntegration.h"
#include "Global.h"
#include "Logger.h"

namespace AIChatSourceBackup {
namespace {
std::wstring Decode(const std::string& value, UINT codePage)
{
    if (value.empty()) return {};
    const int size = MultiByteToWideChar(codePage, 0, value.data(), static_cast<int>(value.size()), nullptr, 0);
    std::wstring result(size, L'\0');
    if (size) MultiByteToWideChar(codePage, 0, value.data(), static_cast<int>(value.size()), result.data(), size);
    return result;
}
}

bool IsEnabled(const AIJsonConfig& config)
{
    const auto value = config.getGlobalValue(kEnabledKey);
    return value.empty() || value == "true" || value == "1";
}

int GetKeepCount(const AIJsonConfig& config)
{
    const auto value = config.getGlobalValue(kKeepCountKey);
    int count = 0;
    const auto parsed = std::from_chars(value.data(), value.data() + value.size(), count);
    return parsed.ec == std::errc{} && parsed.ptr == value.data() + value.size() &&
        count >= 1 && count <= kMaxKeepCount ? count : kDefaultKeepCount;
}

void CaptureBeforeSend(const AIJsonConfig* config, const std::string& requestUtf8)
{
    try {
        if (config && !IsEnabled(*config)) return;
        if (!g_hwnd || GetWindowThreadProcessId(g_hwnd, nullptr) != GetCurrentThreadId()) {
            Logger::Instance().WriteAndIde("AI Backup", "backup skipped: requires IDE main thread"); return;
        }
        UpdateCurrentOpenSourceFile();
        const std::filesystem::path source(Decode(g_nowOpenSourceFilePath, CP_ACP));
        std::filesystem::path output;
        std::string error;
        const bool ok = CreateBackup(source, Decode(requestUtf8, CP_UTF8),
            config ? GetKeepCount(*config) : kDefaultKeepCount,
            [](const std::filesystem::path& path, std::string& error) {
                // IDE 序列化器只接受本地代码页，短路径避免临时目录用户名称丢字。
                const DWORD size = GetShortPathNameW(path.parent_path().c_str(), nullptr, 0);
                std::wstring parent(size, L'\0');
                if (size && GetShortPathNameW(path.parent_path().c_str(), parent.data(), size)) parent.resize(wcslen(parent.c_str()));
                else parent = path.parent_path().wstring();
                const auto localPath = std::filesystem::path(parent) / path.filename();
                BOOL usedDefault = FALSE;
                const auto wide = localPath.wstring();
                const bool utf8 = GetACP() == CP_UTF8;
                const int bytes = WideCharToMultiByte(CP_ACP, utf8 ? 0 : WC_NO_BEST_FIT_CHARS,
                    wide.c_str(), -1, nullptr, 0, nullptr, utf8 ? nullptr : &usedDefault);
                if (!bytes || usedDefault) { error = "snapshot_path_not_representable"; return false; }
                size_t written = 0; std::string trace;
                return EPackagerIntegration::WriteCurrentProjectSnapshot(localPath, written, trace, error);
            }, output, error);
        Logger::Instance().WriteAndIde("AI Backup", ok
            ? "source backup created" + (error.empty() ? std::string() : "; retention warning: " + error)
            : "source backup skipped/failed: " + error);
    } catch (const std::exception& ex) {
        Logger::Instance().WriteAndIde("AI Backup", std::string("source backup failed: ") + ex.what());
    }
}
} // namespace AIChatSourceBackup
