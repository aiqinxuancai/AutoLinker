#include "PathHelper.h"
#include <algorithm>
#include <chrono>
#include <cctype>
#include <filesystem>
#include <fstream>
#include <mutex>
#include <optional>
#include <regex>
#include <string>
#include <vector>

namespace {

constexpr int kDefaultLogRetentionDays = 3;

std::filesystem::path GetAutoLinkerLogDirectoryPathRaw() {
    return GetAutoLinkerDirectoryPath() / "Log";
}

std::string ToLowerAsciiForPath(std::string text) {
    std::transform(text.begin(), text.end(), text.begin(), [](unsigned char ch) {
        return static_cast<char>(std::tolower(ch));
    });
    return text;
}

bool IsProcessScopedLogFileName(const std::string& fileName) {
    const std::filesystem::path path(fileName);
    return ToLowerAsciiForPath(path.extension().string()) == ".log";
}

std::filesystem::path AddCurrentProcessIdToLogFileName(const std::string& fileName) {
    const std::filesystem::path path(fileName);
    if (!IsProcessScopedLogFileName(fileName)) {
        return path;
    }

    const std::string pidSuffix = "_pid" + std::to_string(GetCurrentProcessId());
    return path.parent_path() / (path.stem().string() + pidSuffix + path.extension().string());
}

void CleanupOldAutoLinkerLogFilesInDirectory(const std::filesystem::path& dir, int retentionDays) {
    if (retentionDays <= 0) {
        retentionDays = kDefaultLogRetentionDays;
    }

    std::error_code ec;
    if (!std::filesystem::exists(dir, ec) || !std::filesystem::is_directory(dir, ec)) {
        return;
    }

    const auto cutoff =
        std::filesystem::file_time_type::clock::now() - std::chrono::hours(24 * retentionDays);
    for (std::filesystem::directory_iterator it(dir, ec), end; it != end && !ec; it.increment(ec)) {
        std::error_code fileEc;
        const auto& entry = *it;
        if (!entry.is_regular_file(fileEc) || fileEc) {
            continue;
        }
        if (ToLowerAsciiForPath(entry.path().extension().string()) != ".log") {
            continue;
        }

        const auto lastWriteTime = entry.last_write_time(fileEc);
        if (fileEc || lastWriteTime >= cutoff) {
            continue;
        }

        std::filesystem::remove(entry.path(), fileEc);
    }
}

void EnsureOldLogCleanupOnce(const std::filesystem::path& dir) {
    static std::once_flag s_cleanupOnce;
    std::call_once(s_cleanupOnce, [&dir]() {
        CleanupOldAutoLinkerLogFilesInDirectory(dir, kDefaultLogRetentionDays);
    });
}

} // namespace

std::string GetBasePath() {
    char buffer[MAX_PATH];
    GetModuleFileName(NULL, buffer, MAX_PATH);
    std::string::size_type pos = std::string(buffer).find_last_of("\\/");
    return std::string(buffer).substr(0, pos);
}
std::filesystem::path GetAutoLinkerDirectoryPath() {
    return std::filesystem::path(GetBasePath()) / "AutoLinker";
}
std::filesystem::path GetAutoLinkerLogDirectoryPath() {
    std::filesystem::path dir = GetAutoLinkerLogDirectoryPathRaw();
    std::error_code ec;
    std::filesystem::create_directories(dir, ec);
    EnsureOldLogCleanupOnce(dir);
    return dir;
}
std::filesystem::path GetAutoLinkerLogFilePath(const std::string& fileName) {
    return GetAutoLinkerLogDirectoryPath() / AddCurrentProcessIdToLogFileName(fileName);
}
void CleanupOldAutoLinkerLogFiles(int retentionDays) {
    const std::filesystem::path dir = GetAutoLinkerLogDirectoryPathRaw();
    std::error_code ec;
    std::filesystem::create_directories(dir, ec);
    if (ec) {
        return;
    }
    CleanupOldAutoLinkerLogFilesInDirectory(dir, retentionDays);
}
std::filesystem::path GetAutoLinkerSessionRootDirectoryPath() {
    std::filesystem::path dir = GetAutoLinkerDirectoryPath() / "Session";
    std::error_code ec;
    std::filesystem::create_directories(dir, ec);
    return dir;
}
std::string SanitizePathComponentForStorage(const std::string& text) {
    std::string sanitized;
    sanitized.reserve(text.size());
    for (char ch : text) {
        switch (ch) {
        case '<':
        case '>':
        case ':':
        case '"':
        case '/':
        case '\\':
        case '|':
        case '?':
        case '*':
            sanitized.push_back('_');
            break;
        default:
            if (static_cast<unsigned char>(ch) < 32) {
                sanitized.push_back('_');
            }
            else {
                sanitized.push_back(ch);
            }
            break;
        }
    }
    while (!sanitized.empty() && (sanitized.back() == ' ' || sanitized.back() == '.')) {
        sanitized.pop_back();
    }
    if (sanitized.empty()) {
        return "Unnamed";
    }
    return sanitized;
}
std::string ExtractBetweenDashes(const std::string& text) {
    std::string delimiter = " - ";
    // 找到第一个 " - " 的位置。
    size_t start = text.find(delimiter);
    if (start == std::string::npos) {
        // 没有找到分隔符，返回空字符串。
        return "";
    }
    start += delimiter.length();
    // 从 start 位置开始查找第二个 " - "。
    size_t end = text.find(delimiter, start);
    if (end == std::string::npos) {
        // 没有找到第二个分隔符，返回空字符串。
        return "";
    }
    // 取出两个分隔符之间的字符串。
    return text.substr(start, end - start);
}
/// <summary>
/// 查找字节序列在文件中的偏移。
/// </summary>
/// <param name="filename"></param>
/// <param name="search_bytes"></param>
/// <returns></returns>
std::optional<size_t> FindByteInFile(const std::string& filename, const std::vector<char>& search_bytes) {
    std::ifstream file(filename, std::ios::binary);
    std::vector<char> file_contents((std::istreambuf_iterator<char>(file)), std::istreambuf_iterator<char>());
    auto it = std::search(file_contents.begin(), file_contents.end(), search_bytes.begin(), search_bytes.end());
    if (it != file_contents.end()) {
        return std::distance(file_contents.begin(), it);
    }
    return std::nullopt;
}
/// <summary>
/// 获取 /out: 对应的输出文件名。
/// </summary>
/// <param name="s"></param>
/// <returns></returns>
std::string GetLinkerCommandOutFileName(const std::string& s) {
    std::regex reg("/out:\"([^\"]*)\"");
    std::smatch match;
    if (std::regex_search(s, match, reg) && match.size() > 1) {
        std::string path = match.str(1);
        std::filesystem::path fs_path(path);
        return fs_path.filename().string();
    }
    else {
        return "";
    }
}
// 从命令行文本中提取包含特定目标的完整路径。
std::string ExtractPathFromCommand(const std::string& commandLine, const std::string& target) {
    std::string foundPath;
    size_t pos = commandLine.find(target);
    if (pos != std::string::npos) {
        size_t start = commandLine.rfind('"', pos);
        size_t end = commandLine.find('"', pos + target.length());
        if (start != std::string::npos && end != std::string::npos) {
            foundPath = commandLine.substr(start + 1, end - start - 1);
        }
    }
    return foundPath;
}
std::string GetLinkerCommandKrnlnFileName(const std::string& s) {
    std::string target = "\\static_lib\\krnln_static.lib";
    std::string foundPath = ExtractPathFromCommand(s, target);
    return foundPath;
}
