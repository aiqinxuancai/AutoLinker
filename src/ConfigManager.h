#pragma once
// ConfigManager.h
#ifndef CONFIG_MANAGER_H
#define CONFIG_MANAGER_H

#include <fstream>
#include <string>
#include <map>
#include <filesystem>
#include <mutex>

// 简单 INI 配置管理器，负责 AutoLinker 的键值配置读写。
class ConfigManager {
public:
    ConfigManager();
    std::string getValue(const std::string& key);
    void setValue(const std::string& key, const std::string& value);

private:
    void loadConfig();
    void saveConfig();

    std::filesystem::path configFilePath;
    std::map<std::string, std::string> configData;
    mutable std::mutex configMutex;
};

#endif // CONFIG_MANAGER_H
