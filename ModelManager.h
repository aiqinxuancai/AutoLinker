#pragma once
// ModelManager.h
#ifndef MODEL_MANAGER_H
#define MODEL_MANAGER_H

#include <fstream>
#include <string>
#include <map>
#include <filesystem>

class ModelManager {
public:
    ModelManager();
    std::string getValue(const std::string& key);
    void setValue(const std::string& key, const std::string& value);

private:
    void loadConfig();
    void saveConfig();

    std::filesystem::path configFilePath;
    std::map<std::string, std::string> configData;
};

#endif // MODEL_MANAGER_H