// 独立回归测试：链接真实镜像实现，用桩模拟 IDE 查询源码时同步通知镜像清理。
#include <Windows.h>
#include "../src/WorkspaceMirror.h"
#include "../src/EPackagerIntegration.h"
#include <functional>
#include <fstream>
#include <iostream>
#include <stdexcept>

namespace {
bool sourceChanged = false;
bool sourceAvailable = true;
unsigned sourceNumber = 0;
unsigned resetCount = 0;
unsigned unpackCount = 0;
void Check(bool ok, const char* message)
{
    if (!ok) throw std::runtime_error(message);
}
}

void OutputStringToELog(const std::string&) {}

namespace EPackagerIntegration {
bool GetCurrentSourcePath(std::filesystem::path& path, std::string& error)
{
    // 与 UpdateCurrentOpenSourceFile -> OnCurrentSourceFilePathChanged 的同步回调一致。
    if (sourceChanged) {
        sourceChanged = false;
        ++sourceNumber;
        WorkspaceMirror::ResetAndCleanup();
        ++resetCount;
    }
    error.clear();
    path.clear();
    if (!sourceAvailable) {
        error = "no current source";
        return false;
    }
    path = std::filesystem::temp_directory_path() / ("mirror-reentry-" + std::to_string(sourceNumber) + ".e");
    return true;
}
std::filesystem::path BuildCurrentProjectSnapshotPathForSource(const std::filesystem::path& path)
{
    return path;
}
bool WriteCurrentProjectSnapshot(const std::filesystem::path&, size_t& bytes, std::string& trace, std::string& error)
{
    bytes = 1;
    trace.clear();
    error.clear();
    return true;
}
void CleanupSnapshotRoot(const std::filesystem::path&) {}
bool EnsureToolReady(std::filesystem::path& path, std::string& error)
{
    path = "test-packager.exe";
    error.clear();
    return true;
}
ProcessRunResult RunProcessAndCapture(const std::filesystem::path&, const std::vector<std::wstring>& args,
    const std::filesystem::path&, unsigned long)
{
    ++unpackCount;
    const std::filesystem::path root(args.at(2));
    std::filesystem::create_directories(root / "src");
    std::filesystem::create_directories(root / "project");
    std::ofstream(root / "src" / "page.txt") << "source " << sourceNumber;
    std::ofstream(root / "project" / "_meta.json")
        << R"({"sourceFiles":[{"relativePath":"src/page.txt","logicalName":"page"}]})";
    ProcessRunResult result;
    result.ok = true;
    result.elapsedMs = 1;
    return result;
}
}

int main()
{
    try {
        std::string error;
        std::string relative;
        std::string mode;
        std::filesystem::path path;
        WorkspaceMirror::ProgramItemRef item;
        std::vector<std::string> files;
        const std::vector<std::pair<const char*, std::function<bool()>>> cases = {
            {"ensure", [&] { return WorkspaceMirror::EnsureMirrorFresh(error); }},
            {"prepare", [&] { return WorkspaceMirror::PrepareFileAccess(error); }},
            {"refresh-auto", [&] { return WorkspaceMirror::RefreshMirror(error, &mode); }},
            {"refresh-main", [&] { return WorkspaceMirror::RefreshMirror(error, &mode, WorkspaceMirror::RefreshMode::MainOnly); }},
            {"refresh-full", [&] { return WorkspaceMirror::RefreshMirror(error, &mode, WorkspaceMirror::RefreshMode::Full); }},
            {"update", [&] { return WorkspaceMirror::UpdateMirrorTextFile("src/page.txt", "updated", error); }},
            {"root", [&] { return WorkspaceMirror::GetMirrorRoot(path, error); }},
            {"resolve", [&] { return WorkspaceMirror::ResolveFilePath("src/page.txt", path, relative, error); }},
            {"item", [&] { return WorkspaceMirror::ResolveFileToProgramItem("src/page.txt", item, error); }},
            {"list", [&] { return WorkspaceMirror::ListMirrorFiles(files, error); }}
        };
        for (const auto& [name, run] : cases) {
            sourceAvailable = true;
            Check(WorkspaceMirror::EnsureMirrorFresh(error), "prepare old mirror");
            WorkspaceMirror::FileAccessSnapshot before;
            Check(WorkspaceMirror::GetPreparedFileAccessSnapshot(before, error), "old snapshot");
            const auto previousResets = resetCount;
            sourceChanged = true;
            Check(run(), error.c_str());
            Check(resetCount == previousResets + 1, "source callback did not complete");
            WorkspaceMirror::FileAccessSnapshot after;
            Check(WorkspaceMirror::GetPreparedFileAccessSnapshot(after, error), "new snapshot");
            Check(after.generation > before.generation && after.mirrorRoot != before.mirrorRoot, "stale mirror reused");
            Check(!std::filesystem::exists(before.mirrorRoot), "old mirror not cleaned");
            const auto previousUnpacks = unpackCount;
            Check(WorkspaceMirror::EnsureMirrorFresh(error), "reuse current mirror");
            Check(unpackCount == previousUnpacks, "unchanged source unnecessarily rebuilt");
            sourceAvailable = false;
            sourceChanged = true;
            Check(!run() && error == "no current source", "closed source handling");
            if (std::string(name) == "root") Check(path.empty(), "failed root output not cleared");
            std::cout << "PASS " << name << ": source change and source close\n";
        }
        WorkspaceMirror::ResetAndCleanup();
        std::cout << "PASS all 20 reentry cases\n";
        return 0;
    } catch (const std::exception& ex) {
        std::cerr << "FAIL: " << ex.what() << '\n';
        WorkspaceMirror::ResetAndCleanup();
        return 1;
    }
}
