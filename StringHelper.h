#pragma once
#include <string>

std::string ReplaceSubstring(std::string source, const std::string& toFind, const std::string& toReplace);

std::vector<std::string> ReadFileAndSplitLines(const std::string& filePath);