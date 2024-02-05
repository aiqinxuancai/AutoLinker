#pragma once
#include <string>

std::string ReplaceSubstring(std::string source, const std::string& toFind, const std::string& toReplace);

std::vector<std::string> ReadFileAndSplitLines(const std::string& filePath);


/// <summary>
/// 根据符号分割为两半，如果没符号，则为一个原文本
/// </summary>
/// <param name="input"></param>
/// <param name="delimiter"></param>
/// <returns></returns>
std::vector<std::string> SplitStringTwo(const std::string& input, char delimiter);