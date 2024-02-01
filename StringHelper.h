#pragma once
#include <string>

std::string ReplaceSubstring(std::string source, const std::string& toFind, const std::string& toReplace);

std::vector<std::string> ReadFileAndSplitLines(const std::string& filePath);


/// <summary>
/// ���ݷ��ŷָ�Ϊ���룬���û���ţ���Ϊһ��ԭ�ı�
/// </summary>
/// <param name="input"></param>
/// <param name="delimiter"></param>
/// <returns></returns>
std::vector<std::string> SplitStringTwo(const std::string& input, char delimiter);