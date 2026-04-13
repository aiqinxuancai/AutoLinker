// ConfigManager.cpp
#include "ConfigManager.h"
#include "PathHelper.h"

#include <algorithm>
#include <cctype>

namespace {

// value 中的换行和反斜杠转义（存文件时使用）
static std::string EscapeConfigValue(const std::string& value)
{
    std::string out;
    out.reserve(value.size());
    for (unsigned char c : value) {
        if      (c == '\\') out += "\\\\";
        else if (c == '\r') out += "\\r";
        else if (c == '\n') out += "\\n";
        else out.push_back(static_cast<char>(c));
    }
    return out;
}

// 与 EscapeConfigValue 对应的反转义（读文件时使用）
static std::string UnescapeConfigValue(const std::string& value)
{
    std::string out;
    out.reserve(value.size());
    bool esc = false;
    for (char c : value) {
        if (esc) {
            if      (c == '\\') out.push_back('\\');
            else if (c == 'r')  out.push_back('\r');
            else if (c == 'n')  out.push_back('\n');
            else { out.push_back('\\'); out.push_back(c); }
            esc = false;
        } else if (c == '\\') {
            esc = true;
        } else {
            out.push_back(c);
        }
    }
    if (esc) out.push_back('\\');
    return out;
}

} // namespace

ConfigManager::ConfigManager()
{
    std::filesystem::path currentPath = GetBasePath();
    std::filesystem::path autoLinkerPath = currentPath / "AutoLinker";
    if (!std::filesystem::exists(autoLinkerPath)) {
        std::filesystem::create_directory(autoLinkerPath);
    }
    configFilePath = autoLinkerPath / "AutoLinker.ini";
    loadConfig();
}

std::string ConfigManager::getValue(const std::string& key)
{
    std::string k = key;
    std::transform(k.begin(), k.end(), k.begin(),
        [](unsigned char c) { return std::tolower(c); });
    const auto it = configData.find(k);
    return it != configData.end() ? it->second : std::string();
}

void ConfigManager::setValue(const std::string& key, const std::string& value)
{
    std::string k = key;
    std::transform(k.begin(), k.end(), k.begin(),
        [](unsigned char c) { return std::tolower(c); });
    configData[k] = value;
    saveConfig();
}

void ConfigManager::loadConfig()
{
    std::ifstream configFile(configFilePath);
    std::string line;
    while (std::getline(configFile, line)) {
        // 去掉 Windows 换行残留的 \r
        if (!line.empty() && line.back() == '\r') line.pop_back();
        const size_t pos = line.find('=');
        if (pos != std::string::npos) {
            std::string k = line.substr(0, pos);
            std::string v = UnescapeConfigValue(line.substr(pos + 1));
            configData[k] = std::move(v);
        }
    }
}

void ConfigManager::saveConfig()
{
    std::ofstream configFile(configFilePath);
    for (const auto& [key, value] : configData) {
        configFile << key << '=' << EscapeConfigValue(value) << '\n';
    }
}