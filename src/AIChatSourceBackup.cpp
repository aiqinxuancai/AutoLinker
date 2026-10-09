#include "AIChatSourceBackup.h"

#include <Windows.h>
#include <charconv>
#include "AIJsonConfig.h"
#include "EPackagerIntegration.h"
#include "Global.h"
#include "AutoLinkerInternal.h"
#include "AIChatFeature.h"
#include "AIChatToolRegistry.h"
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

namespace {
bool capturing = false;
int ReadCount(const AIJsonConfig& config, const char* key, int fallback, int maximum)
{
    const auto value = config.getGlobalValue(key);
    int number = 0;
    const auto result = std::from_chars(value.data(), value.data() + value.size(), number);
    return result.ec == std::errc{} && result.ptr == value.data() + value.size() && number >= 1 && number <= maximum ? number : fallback;
}
}
bool IsEnabled(const AIJsonConfig& config, Kind kind)
{
    if (kind == Kind::Chat) return IsEnabled(config);
    const auto value = config.getGlobalValue(kind == Kind::Timed ? kTimedEnabledKey : kToolEnabledKey);
    return value == "true" || value == "1";
}
int GetKeepCount(const AIJsonConfig& config, Kind kind)
{
    if (kind == Kind::Chat) return GetKeepCount(config);
    return ReadCount(config, kind == Kind::Timed ? kTimedKeepKey : kToolKeepKey, 20, kMaxKeepCount);
}
int GetIntervalMinutes(const AIJsonConfig& config) { return ReadCount(config, kIntervalKey, 5, 1440); }

static void Capture(const AIJsonConfig* config, std::wstring_view request, Kind kind)
{
    try {
        if (config && !IsEnabled(*config, kind)) return;
        if (capturing || g_preDebugging || g_preCompiling || AIChatFeature::IsProjectRestoreActive()) return;
        struct Guard { Guard() { capturing = true; } ~Guard() { capturing = false; } } guard;
        if (!g_hwnd || GetWindowThreadProcessId(g_hwnd, nullptr) != GetCurrentThreadId()) {
            Logger::Instance().WriteAndIde("AI Backup", "backup skipped: requires IDE main thread"); return;
        }
        UpdateCurrentOpenSourceFile();
        const std::filesystem::path source(Decode(g_nowOpenSourceFilePath, CP_ACP));
        std::filesystem::path output;
        std::string error;
        const bool ok = CreateBackup(source, request,
            config ? GetKeepCount(*config, kind) : kDefaultKeepCount,
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
            }, output, error, true, kind);
        Logger::Instance().WriteAndIde("AI Backup", ok
            ? "source backup created" + (error.empty() ? std::string() : "; retention warning: " + error)
            : "source backup skipped/failed: " + error);
    } catch (const std::exception& ex) {
        Logger::Instance().WriteAndIde("AI Backup", std::string("source backup failed: ") + ex.what());
    }
}
void CaptureBeforeSend(const AIJsonConfig* config, const std::string& requestUtf8)
{
    Capture(config, Decode(requestUtf8, CP_UTF8), Kind::Chat);
}
void CaptureBeforeTool(const AIJsonConfig* config, const std::string& toolName)
{
    if (config && AIChatToolRegistry::ModifiesProject(toolName)) Capture(config, Decode(toolName, CP_UTF8), Kind::ToolWrite);
}
void Tick(const AIJsonConfig* config, bool busy)
{
    static ULONGLONG due = 0;
    static std::string project;
    static int previousInterval = 0;
    if (!config || !IsEnabled(*config, Kind::Timed)) { due = 0; project.clear(); return; }
    if (busy || capturing || g_preDebugging || g_preCompiling || AIChatFeature::IsProjectRestoreActive()) return;
    UpdateCurrentOpenSourceFile();
    const auto now = GetTickCount64();
    const int interval = GetIntervalMinutes(*config);
    if (!due || project != g_nowOpenSourceFilePath || previousInterval != interval) {
        project = g_nowOpenSourceFilePath; previousInterval = interval;
        due = now + static_cast<ULONGLONG>(interval) * 60000;
    }
    if (now < due || project.empty()) return;
    due = now + static_cast<ULONGLONG>(interval) * 60000;
    Capture(config, L"定时快照", Kind::Timed);
}
} // namespace AIChatSourceBackup
