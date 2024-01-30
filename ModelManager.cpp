// ModelManager.cpp
#include "ModelManager.h"
#include "PathHelper.h"

ModelManager::ModelManager() {
    // 获取当前进程的路径
    std::filesystem::path currentPath = GetBasePath();
    // 创建AutoLoader目录
    std::filesystem::path autoLoaderPath = currentPath / "AutoLoader";
    if (!std::filesystem::exists(autoLoaderPath)) {
        std::filesystem::create_directory(autoLoaderPath);
    }
    // 创建并打开配置文件
    configFilePath = autoLoaderPath / "ModelManager.ini";
    loadConfig();
}

std::string ModelManager::getValue(const std::string& key) {

    std::string k = key;

    //std::transform(k.begin(), k.end(), k.begin(),
    //    [](unsigned char c) { return std::tolower(c); });

    if (configData.contains(k)) {
        return configData[k];
    }
    return std::string();
    
}

void ModelManager::setValue(const std::string& key, const std::string& value) {
    //将key转为路径?
    //std::filesystem::path p = key;

    std::string k = key;

    //std::transform(k.begin(), k.end(), k.begin(),
    //    [](unsigned char c) { return std::tolower(c); });

    configData[k] = value;
    saveConfig();
}

std::string ModelManager::getKeyFromValue(const std::string& value) {
    for (const auto& kv : configData) {
        if (kv.second == value) {
            return kv.first; // 找到值后返回对应的键
        }
    }
    return std::string(); // 如果没有找到，返回空字符串
}

void ModelManager::loadConfig() {
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

void ModelManager::saveConfig() {
    std::ofstream configFile(configFilePath);
    for (const auto& [key, value] : configData) {
        configFile << key << '=' << value << '\n';
    }
}