#pragma once

#include <vector>
#include <windows.h>
#include <psapi.h>
#include <algorithm>
#include <sstream>
#include <iomanip>


int FindSelfModelMemory(const std::string& pattern_str);