
#include "PathHelper.h"
#include <optional>
#include <vector>
#include <regex>
#include <fstream>
#include <algorithm>
#include <regex>
#include <filesystem>


std::string GetBasePath() {
    char buffer[MAX_PATH];
    GetModuleFileName(NULL, buffer, MAX_PATH);
    std::string::size_type pos = std::string(buffer).find_last_of("\\/");
    return std::string(buffer).substr(0, pos);
}

std::string ExtractBetweenDashes(const std::string& text) {
    std::string delimiter = " - ";

    // 找到第一个" - "的位置
    size_t start = text.find(delimiter);
    if (start == std::string::npos) {
        // 没有找到" - "，返回空字符串
        return "";
    }
    start += delimiter.length(); // 跳过" - "，开始于第一个" - "之后的字符

    // 从start位置开始，找到下一个" - "的位置
    size_t end = text.find(delimiter, start);
    if (end == std::string::npos) {
        // 没有找到第二个" - "，返回空字符串
        return "";
    }

    // 取出两个" - "之间的子字符串
    return text.substr(start, end - start);
}

/// <summary>
/// 读取
/// </summary>
/// <param name="filename"></param>
/// <param name="search_bytes"></param>
/// <returns></returns>
std::optional<size_t> FindByteInFile(const std::string& filename, const std::vector<char>& search_bytes) {
    // 打开文件并读取内容
    std::ifstream file(filename, std::ios::binary);
    std::vector<char> file_contents((std::istreambuf_iterator<char>(file)), std::istreambuf_iterator<char>());

    // 搜索字节序列
    auto it = std::search(file_contents.begin(), file_contents.end(), search_bytes.begin(), search_bytes.end());

    if (it != file_contents.end()) {
        // 如果找到了，返回位置
        return std::distance(file_contents.begin(), it);
    }
    else {
        // 如果没有找到，返回 std::nullopt
        return std::nullopt;
    }
}




/// <summary>
/// 获取out的文件名
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



// 函数：从命令行文本中提取包含特定特征的路径
std::string ExtractPathFromCommand(const std::string& commandLine, const std::string& target) {
    std::string foundPath;
    size_t pos = commandLine.find(target);
    if (pos != std::string::npos) {
        // 从特征字符串的位置向前查找第一个双引号
        size_t start = commandLine.rfind('"', pos);
        // 从特征字符串的位置向后查找第一个双引号
        size_t end = commandLine.find('"', pos + target.length());

        // 如果找到了双引号，提取路径
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

