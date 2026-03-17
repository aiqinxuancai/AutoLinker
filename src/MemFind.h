#pragma once

#include <vector>
#include <windows.h>
#include <psapi.h>
#include <algorithm>
#include <sstream>
#include <iomanip>


int FindSelfModelMemory(const std::string& pattern_str);
std::vector<int> FindSelfModelMemoryAll(const std::string& pattern_str);
int FindSelfModelMemoryUnique(const std::string& pattern_str);
std::vector<size_t> FindPatternOffsets(const byte* data, size_t dataSize, const std::string& patternStr);
