// 工具目录唯一性防御：相同定义去重，歧义名称整组拒绝，保持其余工具原始顺序。
#pragma once

#include <string>
#include <unordered_map>
#include <vector>
#include "..\\thirdparty\\json.hpp"

namespace AIToolCatalogGuard {
struct Result {
    nlohmann::json catalog = nlohmann::json::array();
    std::vector<size_t> indices;
    std::vector<std::string> duplicateNames;
    std::vector<std::string> conflictNames;
    size_t invalidCount = 0;

    bool Changed() const { return invalidCount || !duplicateNames.empty() || !conflictNames.empty(); }
};

inline Result Filter(const nlohmann::json& catalog)
{
    Result result;
    if (!catalog.is_array()) { result.invalidCount = 1; return result; }
    struct Entry { size_t first; bool conflict = false; bool duplicate = false; };
    std::unordered_map<std::string, Entry> names;
    for (size_t index = 0; index < catalog.size(); ++index) {
        const auto& item = catalog[index];
        if (!item.is_object() || !item.contains("name") || !item["name"].is_string() ||
            item["name"].get_ref<const std::string&>().find_first_not_of(" \t\r\n") == std::string::npos) {
            ++result.invalidCount;
            continue;
        }
        const std::string name = item["name"].get<std::string>();
        auto [it, inserted] = names.try_emplace(name, Entry{index});
        if (inserted) continue;
        auto& entry = it->second;
        if (catalog[entry.first] != item) {
            if (!entry.conflict) result.conflictNames.push_back(name);
            entry.conflict = true;
        } else if (!entry.duplicate) {
            result.duplicateNames.push_back(name);
            entry.duplicate = true;
        }
    }
    for (size_t index = 0; index < catalog.size(); ++index) {
        const auto& item = catalog[index];
        if (!item.is_object() || !item.contains("name") || !item["name"].is_string()) continue;
        const auto it = names.find(item["name"].get<std::string>());
        if (it == names.end() || it->second.conflict || it->second.first != index) continue;
        result.catalog.push_back(item);
        result.indices.push_back(index);
    }
    return result;
}
}
