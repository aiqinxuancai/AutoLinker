#include <iostream>
#include <string>
#include <vector>
#include <fstream>
#include <sstream>
#include <string>
#include <vector>
#include <filesystem> 


std::string ReplaceSubstring(std::string source, const std::string& toFind, const std::string& toReplace) {
    size_t pos = 0;
    while ((pos = source.find(toFind, pos)) != std::string::npos) {
        source.replace(pos, toFind.length(), toReplace);
        pos += toReplace.length();
    }
    return source;
}


std::vector<std::string> ReadFileAndSplitLines(const std::string& filePath) {
    // 检查文件是否存在
    if (!std::filesystem::exists(filePath)) {
        return {}; // 如果文件不存在，返回空的vector
    }

    std::ifstream file(filePath);
    if (!file.is_open()) {
        return {}; // 如果文件无法打开，返回空的vector
    }

    std::stringstream buffer;
    buffer << file.rdbuf(); // 读取文件到字符串流

    std::string content = buffer.str();

    // 替换所有 \r\n 为 \n
    size_t pos = 0;
    while ((pos = content.find("\r\n", pos)) != std::string::npos) {
        content.replace(pos, 2, "\n");
        pos += 1;
    }

    // 按行分割
    std::vector<std::string> lines;
    std::istringstream stream(content);
    std::string line;
    while (std::getline(stream, line)) {
        lines.push_back(line);
    }

    return lines;
}

/// <summary>
/// 根据符号分割为两半，如果没找到'='，则为一个原文本
/// </summary>
/// <param name="input"></param>
/// <param name="delimiter"></param>
/// <returns></returns>
std::vector<std::string> SplitStringTwo(const std::string& input, char delimiter) {
    std::vector<std::string> result;
    size_t pos = input.find(delimiter);

    if (pos != std::string::npos) {
        // 如果找到了分隔符
        std::string first = input.substr(0, pos);
        std::string second = input.substr(pos + 1);

        // 添加两部分，即使第二部分为空也添加
        result.push_back(first);
        result.push_back(second);
    }
    else {
        // 如果没有找到分隔符，返回原字符串
        result.push_back(input);
    }

    return result;
}