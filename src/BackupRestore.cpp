#include "BackupRestore.h"
#include <Windows.h>
#include <algorithm>
#include <fstream>
#include <atomic>
#include "AIChatSourceBackup.h"
#include "AIChatFeature.h"
#include "AutoLinkerInternal.h"
#include "EideProjectRestore.h"
#include "EideProjectBinarySerializer.h"
#include "EPackagerIntegration.h"
#include "PageCodeCacheManager.h"
#include "Global.h"
#include "Logger.h"
#include "WindowHelper.h"

namespace BackupRestore {
namespace {
struct Staging {
    std::filesystem::path directory;
    std::filesystem::path selected, rollback;
    ~Staging() {
        std::error_code ec;
        if (!selected.empty()) std::filesystem::remove(selected, ec);
        if (!rollback.empty()) std::filesystem::remove(rollback, ec);
        if (!directory.empty()) std::filesystem::remove_all(directory, ec);
    }
};
struct Transaction {
    ~Transaction() {
        e571::ProjectBinarySerializer::Instance().ClearVerifiedSerializerContext();
        PageCodeCacheManager::Instance().Clear();
        AIChatFeature::EndProjectRestore();
    }
};
std::string LocalPath(const std::filesystem::path& path)
{
    std::wstring wide = path.wstring();
    const DWORD size = GetShortPathNameW(path.c_str(), nullptr, 0);
    if (size) {
        std::wstring shortPath(size, 0);
        if (GetShortPathNameW(path.c_str(), shortPath.data(), size)) { shortPath.resize(wcslen(shortPath.c_str())); wide = shortPath; }
    }
    BOOL fallback = FALSE;
    const bool utf8 = GetACP() == CP_UTF8;
    const DWORD flags = utf8 ? 0 : WC_NO_BEST_FIT_CHARS;
    const int count = WideCharToMultiByte(CP_ACP, flags, wide.c_str(), -1, nullptr, 0, nullptr, utf8 ? nullptr : &fallback);
    if (!count || fallback) throw std::runtime_error("snapshot_path_not_representable");
    std::string local(count, 0);
    WideCharToMultiByte(CP_ACP, flags, wide.c_str(), -1, local.data(), count, nullptr, utf8 ? nullptr : &fallback);
    local.pop_back();
    return local;
}
bool WriteSnapshot(const std::filesystem::path& path, std::string& error)
{
    const auto local = std::filesystem::path(LocalPath(path.parent_path())) / path.filename();
    size_t bytes = 0; std::string trace;
    return EPackagerIntegration::WriteCurrentProjectSnapshot(local, bytes, trace, error);
}
bool HasMagic(const std::filesystem::path& path)
{
    std::ifstream input(path, std::ios::binary);
    char magic[8] = {}; input.read(magic, sizeof(magic));
    return input && memcmp(magic, "CNWTEPRG", 8) == 0;
}
}
std::filesystem::path CurrentSource()
{
    UpdateCurrentOpenSourceFile();
    // 事务期间运行时不发布中间路径，但最终校验必须读取真实 IDE 标题。
    const auto local = AIChatFeature::IsProjectRestoreActive()
        ? GetSourceFilePath() : g_nowOpenSourceFilePath;
    const int size = MultiByteToWideChar(CP_ACP, 0, local.data(), static_cast<int>(local.size()), nullptr, 0);
    std::wstring wide(size, 0);
    if (size) MultiByteToWideChar(CP_ACP, 0, local.data(), static_cast<int>(local.size()), wide.data(), size);
    return std::filesystem::path(wide).lexically_normal();
}
bool Restore(const std::filesystem::path& expectedSource, const std::filesystem::path& backup,
    std::filesystem::path& safetyBackup, std::wstring& message)
{
    safetyBackup.clear();
    if (!EideProjectRestore::IsSupported()) { message = L"无法恢复，此功能需要 e5.95 版本。"; return false; }
    if (!g_hwnd || GetCurrentThreadId() != GetWindowThreadProcessId(g_hwnd, nullptr)) { message = L"恢复必须在 IDE 主线程执行。"; return false; }
    if (g_preDebugging || g_preCompiling) { message = L"请先结束调试或编译后再恢复。"; return false; }
    try {
        const auto source = CurrentSource();
        if (source.empty() || source != expectedSource || !std::filesystem::is_regular_file(source)) {
            message = L"当前工程已改变或尚未保存，请刷新列表。"; return false;
        }
        if (_wcsicmp(source.extension().c_str(), L".e") != 0) { message = L"目前仅支持 .e 工程恢复。"; return false; }
        const auto entries = AIChatSourceBackup::ListBackups(source);
        if (std::find(entries.begin(), entries.end(), backup) == entries.end()) { message = L"所选备份不存在或不属于当前工程。"; return false; }
        EideProjectRestore::Identity identity;
        std::string identityError;
        if (!EideProjectRestore::CaptureIdentity(identity, identityError)) {
            message = L"无法读取当前工程身份，未执行恢复。";
            Logger::Instance().WriteAndIde("Backup Restore", identityError);
            return false;
        }
        Staging staging;
        static std::atomic<unsigned long> serial{0};
        do {
            staging.directory = std::filesystem::temp_directory_path() /
                (L"AutoLinkerRestore-" + std::to_wstring(GetCurrentProcessId()) + L"-" + std::to_wstring(GetTickCount64()) + L"-" + std::to_wstring(++serial));
        } while (!std::filesystem::create_directory(staging.directory));
        // 加载器以输入文件目录解析相对模块路径，暂存文件必须与原工程同目录。
        const auto selected = source.parent_path() / (staging.directory.filename().wstring() + L"-selected.e");
        std::filesystem::copy_file(backup, selected);
        staging.selected = selected;
        if (!HasMagic(selected)) { message = L"备份不是有效的完整 .e 快照，未执行恢复。"; return false; }
        const auto selectedLocal = LocalPath(selected);
        if (!AIChatFeature::BeginProjectRestore()) { message = L"AI 或 MCP 任务正在执行，请等待完成后恢复。"; return false; }
        Transaction transaction;
        std::string error;
        const int keep = AIChatSourceBackup::GetKeepCount(g_aiJsonConfig);
        // 无论自动备份是否开启，恢复前都要保存；此时不清理旧文件。
        if (!AIChatSourceBackup::CreateBackup(source, L"恢复前安全备份", keep, WriteSnapshot, safetyBackup, error, false)) {
            message = L"恢复前安全备份失败，未执行恢复。";
            Logger::Instance().WriteAndIde("Backup Restore", error); return false;
        }
        const auto rollback = source.parent_path() / (staging.directory.filename().wstring() + L"-rollback.e");
        std::filesystem::copy_file(safetyBackup, rollback);
        staging.rollback = rollback;
        const auto rollbackLocal = LocalPath(rollback);
        bool ok = false;
        try { ok = EideProjectRestore::Reload(selectedLocal, identity, error) && CurrentSource() == source; }
        catch (...) { error = "reload_exception"; }
        if (!ok) {
            if (error.empty()) error = "source_path_changed_after_reload";
            Logger::Instance().WriteAndIde("Backup Restore", "restore failed: " + error);
            bool recovered = false;
            try { recovered = EideProjectRestore::Reload(rollbackLocal, identity, error) && CurrentSource() == source; } catch (...) {}
            message = recovered ? L"恢复失败，已回滚到恢复前内容；安全备份已保留。" : L"恢复和自动回滚均失败，请使用恢复前安全备份重新打开工程。";
            message += std::wstring(1, 10) + safetyBackup.wstring();
            return false;
        }
        message = L"恢复成功。工程路径不变，内容已标记为未保存；请确认后自行保存。";
        message += std::wstring(1, 10) + L"恢复前安全备份：" + safetyBackup.wstring();
        try { AIChatSourceBackup::PruneBackups(source, safetyBackup, keep); }
        catch (...) { message += L"（旧备份清理失败，已保留现有文件。）"; }
        Logger::Instance().WriteAndIde("Backup Restore", "restored binary project; original path retained; safety snapshot preserved");
        return true;
    } catch (const std::exception& ex) {
        Logger::Instance().WriteAndIde("Backup Restore", ex.what());
        message = L"恢复失败：文件访问或快照操作出错。";
        if (!safetyBackup.empty()) message += std::wstring(1, 10) + L"安全备份：" + safetyBackup.wstring();
        return false;
    }
}
}

