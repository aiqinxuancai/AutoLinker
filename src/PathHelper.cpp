#include "PathHelper.h"
#include <algorithm>
#include <filesystem>
#include <fstream>
#include <optional>
#include <regex>
#include <vector>
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
    std::filesystem::path dir = GetAutoLinkerDirectoryPath() / "Log";
    std::error_code ec;
    std::filesystem::create_directories(dir, ec);
    return dir;
}
std::filesystem::path GetAutoLinkerLogFilePath(const std::string& fileName) {
    return GetAutoLinkerLogDirectoryPath() / fileName;
}
std::string ExtractBetweenDashes(const std::string& text) {
    std::string delimiter = " - ";
    // 鎵惧埌绗竴涓? - "鐨勪綅缃?
    size_t start = text.find(delimiter);
    if (start == std::string::npos) {
        // 娌℃湁鎵惧埌" - "锛岃繑鍥炵┖瀛楃涓?
        return "";
    }
    start += delimiter.length();
    // 浠?start 浣嶇疆寮€濮嬫壘鍒扮浜屼釜" - "鐨勪綅缃?
    size_t end = text.find(delimiter, start);
    if (end == std::string::npos) {
        // 娌℃湁鎵惧埌绗簩涓? - "锛岃繑鍥炵┖瀛楃涓?
        return "";
    }
    // 鍙栧嚭涓や釜" - "涔嬮棿鐨勫瓧绗︿覆
    return text.substr(start, end - start);
}
/// <summary>
/// 鏌ユ壘瀛楄妭搴忓垪鍦ㄦ枃浠朵腑鐨勫亸绉汇€?
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
/// 鑾峰彇 /out: 瀵瑰簲鐨勮緭鍑烘枃浠跺悕銆?
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
// 浠庡懡浠よ鏂囨湰涓彁鍙栧寘鍚壒瀹氱洰鏍囩殑瀹屾暣璺緞銆?
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