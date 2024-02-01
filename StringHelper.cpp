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