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
            return static_cast<int>(i + 1); // 返回匹配的起始索引
        }
    }
    return -1; // 没有找到匹配项
}


int FindSelfModelMemory(const std::string& pattern_str) {
    HMODULE h = GetModuleHandle(NULL);
    if (h == NULL) {
        OutputStringToELog(std::format("获取内存失败"));
        return 0;
    }

    std::vector<byte> module_bytes = GetModuleBytes(h);

    auto address = findPattern(module_bytes, pattern_str);
    if (address != -1) {
        address -= 1;
        return reinterpret_cast<int>(h) + address;
    }
    return 0;
}
