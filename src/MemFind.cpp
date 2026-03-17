#include "MemFind.h"
#include <format>
#include "Global.h"
#include <optional>

std::vector<byte> GetModuleBytes(HMODULE h) {
    MODULEINFO module_info;
    GetModuleInformation(GetCurrentProcess(), h, &module_info, sizeof(module_info));

    std::vector<byte> bytes(module_info.SizeOfImage);
    CopyMemory(bytes.data(), h, bytes.size());

    return bytes;
}

std::optional<byte> hexStringToByte(const std::string& hex) {
    if (hex == "??") return std::nullopt;
    int value;
    std::stringstream ss;
    ss << std::hex << hex;
    ss >> value;
    return static_cast<byte>(value);
}

std::vector<std::optional<byte>> parsePattern(const std::string& patternStr) {
    std::vector<std::optional<byte>> pattern;
    std::istringstream iss(patternStr);
    std::string byteStr;
    while (iss >> byteStr) {
        pattern.push_back(hexStringToByte(byteStr));
    }
    return pattern;
}

int findPattern(const std::vector<byte>& data, const std::string& patternStr) {
    auto pattern = parsePattern(patternStr);
    for (size_t i = 0; i <= data.size() - pattern.size(); ++i) {
        bool match = true;
        for (size_t j = 0; j < pattern.size(); ++j) {
            if (pattern[j].has_value() && data[i + j] != pattern[j].value()) {
                match = false;
                break;
            }
        }
        if (match) {
            return static_cast<int>(i + 1);
        }
    }
    return -1;
}

std::vector<size_t> FindPatternOffsets(const byte* data, size_t dataSize, const std::string& patternStr) {
    std::vector<size_t> offsets;
    if (data == nullptr || dataSize == 0) {
        return offsets;
    }

    const auto pattern = parsePattern(patternStr);
    if (pattern.empty() || dataSize < pattern.size()) {
        return offsets;
    }

    for (size_t i = 0; i <= dataSize - pattern.size(); ++i) {
        bool match = true;
        for (size_t j = 0; j < pattern.size(); ++j) {
            if (pattern[j].has_value() && data[i + j] != pattern[j].value()) {
                match = false;
                break;
            }
        }
        if (match) {
            offsets.push_back(i);
        }
    }
    return offsets;
}

std::vector<int> FindSelfModelMemoryAll(const std::string& pattern_str) {
    std::vector<int> matches;
    HMODULE h = GetModuleHandle(NULL);
    if (h == NULL) {
        OutputStringToELog(std::format("get module memory failed"));
        return matches;
    }

    const std::vector<byte> module_bytes = GetModuleBytes(h);
    const auto offsets = FindPatternOffsets(module_bytes.data(), module_bytes.size(), pattern_str);
    matches.reserve(offsets.size());
    for (const size_t offset : offsets) {
        matches.push_back(reinterpret_cast<int>(h) + static_cast<int>(offset));
    }
    return matches;
}

int FindSelfModelMemoryUnique(const std::string& pattern_str) {
    const auto matches = FindSelfModelMemoryAll(pattern_str);
    return matches.size() == 1 ? matches.front() : 0;
}

int FindSelfModelMemory(const std::string& pattern_str) {
    const auto matches = FindSelfModelMemoryAll(pattern_str);
    return matches.empty() ? 0 : matches.front();
}
