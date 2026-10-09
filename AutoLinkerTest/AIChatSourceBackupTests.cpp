#include "../src/AIChatSourceBackup.h"
#include <Windows.h>
#include <fstream>
#include <iostream>
#include <stdexcept>

using namespace AIChatSourceBackup;
void Check(bool ok, const char* message) { if (!ok) throw std::runtime_error(message); }
std::string Read(const std::filesystem::path& path) { std::ifstream f(path, std::ios::binary); return {std::istreambuf_iterator<char>(f), {}}; }
int main()
{
    const auto root = std::filesystem::temp_directory_path() / (L"AutoLinkerBackupTest-" + std::to_wstring(GetCurrentProcessId()));
    try {
        std::filesystem::create_directories(root);
        const auto source = root / L"测试工程.e";
        { std::ofstream f(source); f << "original"; }
        Check(BuildRequestLabel(L"一二三四五六七八九十一二三四五六七八九十") == L"一二三四五六七八九十一二三四五六七八九十", "twenty characters");
        Check(BuildRequestLabel(L"一二三四五六七八九十一二三四五六七八九十一") == L"一二三四五六七八九十一二三四五六七八九十...", "Chinese truncation");
        const std::wstring emoji = std::wstring(1, static_cast<wchar_t>(0xD83D)) + static_cast<wchar_t>(0xDE00);
        Check(BuildRequestLabel(L"一二三四五六七八九十一二三四五六七八九" + emoji + L"十") == L"一二三四五六七八九十一二三四五六七八九" + emoji + L"...", "surrogate boundary");
        const auto invalid = BuildRequestLabel(L"<>:\"/\\|?*");
        Check(invalid == std::wstring(9, L'_'), "invalid filename characters");
        Check(BuildRequestLabel(L"") == L"图片消息", "attachment fallback");
        std::filesystem::path output; std::string error;
        const auto writer = [](const auto& p, auto&) { std::ofstream f(p, std::ios::binary); f << "unsaved state"; return f.good(); };
        Check(CreateBackup(source, emoji + L"修改", 10, writer, output, error), "Unicode publication");
        Check(output.parent_path().filename() == L"测试工程.e.BAK", "backup directory");
        Check(output.filename().wstring().find(emoji) != std::wstring::npos, "Unicode filename");
        Check(Read(output) == "unsaved state" && Read(source) == "original", "source untouched and memory state copied");
        const auto directory = output.parent_path();
        const auto unrelated = directory / L"manual.e";
        { std::ofstream f(unrelated); f << "keep"; }
        std::filesystem::path previous;
        for (int i = 0; i < 15; ++i) {
            Check(CreateBackup(source, L"重复需求", 10, writer, output, error) && error.empty(), "repeated backup");
            Check(previous != output, "collision does not overwrite");
            previous = output;
        }
        auto count = [&] { size_t n = 0; for (auto& e : std::filesystem::directory_iterator(directory)) if (e.is_regular_file()) ++n; return n; };
        Check(count() == 11 && Read(unrelated) == "keep", "retain ten and preserve unrelated file");
        Check(!CreateBackup(source, L"失败", 1, [](const auto& p, auto& e) { std::ofstream f(p); f << "partial"; e = "simulated failure"; return false; }, output, error), "snapshot failure");
        Check(count() == 11, "failure does not prune existing backups");
        Check(CreateBackup(source, L"保留一份", 1, writer, output, error) && count() == 2, "configurable retention");
        Check(!CreateBackup(root / L"missing.e", L"test", 10, writer, output, error), "unsaved source skipped");
        Check(Read(source) == "original", "original preserved");
        std::filesystem::remove_all(root);
        std::cout << "PASS: Unicode, illegal characters, snapshot contents, collisions, retention, failure cleanup, original preservation" << std::endl;
        return 0;
    } catch (const std::exception& ex) {
        std::cerr << ex.what() << std::endl;
        std::filesystem::remove_all(root); return 1;
    }
}
