// ConfigManager.cpp
#include "ConfigManager.h"
#include "PathHelper.h"

ConfigManager::ConfigManager() {
    // 获取当前进程的路径
    std::filesystem::path currentPath = GetBasePath();
    // 创建AutoLinker目录
    std::filesystem::path autoLinkerPath = currentPath / "AutoLinker";
    if (!std::filesystem::exists(autoLinkerPath)) {
        std::filesystem::create_directory(autoLinkerPath);
    }
    // 创建并打开配置文件
    configFilePath = autoLinkerPath / "AutoLinker.ini";
    loadConfig();
}

std::string ConfigManager::getValue(const std::string& key) {

    std::string k = key;

    std::transform(k.begin(), k.end(), k.begin(),
        [](unsigned char c) { return std::tolower(c); });

    if (configData.contains(k)) {
        return configData[k];
    }
    return std::string();
    
}

void ConfigManager::setValue(const std::string& key, const std::string& value) {
    //将key转为路径?
    //std::filesystem::path p = key;

    std::string k = key;

    std::transform(k.begin(), k.end(), k.begin(),
        [](unsigned char c) { return std::tolower(c); });

    configData[k] = value;
    saveConfig();
}

void ConfigManager::loadConfig() {
    std::ifstream configFile(configFilePath);
    std::string line;
    while (std::getline(configFile, line)) {
        size_t pos = line.find('=');
        if (pos != std::string::npos) {
            std::string key = line.substr(0, pos);
            std::string value = line.substr(pos + 1);
            configData[key] = value;
        }
    }
}

void ConfigManager::saveConfig() {
    std::ofstream configFile(configFilePath);
    for (const auto& [key, value] : configData) {
        configFile << key << '=' << value << '\n';
    }
}