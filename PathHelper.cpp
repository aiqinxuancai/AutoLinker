#include "PathHelper.h"


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