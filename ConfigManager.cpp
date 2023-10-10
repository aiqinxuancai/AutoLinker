// ConfigManager.cpp
#include "ConfigManager.h"
#include "PathHelper.h"

ConfigManager::ConfigManager() {
    // ��ȡ��ǰ���̵�·��
    std::filesystem::path currentPath = GetBasePath();
    // ����AutoLoaderĿ¼
    std::filesystem::path autoLoaderPath = currentPath / "AutoLoader";
    if (!std::filesystem::exists(autoLoaderPath)) {
        std::filesystem::create_directory(autoLoaderPath);
    }
    // �������������ļ�
    configFilePath = autoLoaderPath / "AutoLoader.ini";
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
    //��keyתΪ·��?
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