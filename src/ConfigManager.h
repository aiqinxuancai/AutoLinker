#pragma once
// ConfigManager.h
#ifndef CONFIG_MANAGER_H
#define CONFIG_MANAGER_H

#include <fstream>
#include <string>
#include <map>
#include <filesystem>

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
};

#endif // CONFIG_MANAGER_H