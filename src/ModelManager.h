#pragma once
// ModelManager.h
#ifndef MODEL_MANAGER_H
#define MODEL_MANAGER_H

#include <filesystem>
#include <map>
#include <string>

// 管理 ModelManager.ini 中的 EC 模块动静态切换关系。
class ModelManager {
public:
    ModelManager();
    // 根据动态调试版 ec 文件名查找静态编译版 ec 文件名。
    std::string getValue(const std::string& key);
    // 设置单条切换关系并立即保存。
    void setValue(const std::string& key, const std::string& value);
    // 根据静态编译版 ec 文件名反查动态调试版 ec 文件名。
    std::string getKeyFromValue(const std::string& value);
    // 返回全部切换关系，字符串使用当前系统本地编码。
    std::map<std::string, std::string> getAllValues() const;
    // 替换全部切换关系并保存到 ModelManager.ini。
    bool replaceAllValues(const std::map<std::string, std::string>& values, std::string* errorMessage = nullptr);
    // 返回配置文件路径。
    const std::filesystem::path& getConfigFilePath() const;

private:
    void loadConfig();
    bool saveConfig(std::string* errorMessage = nullptr) const;

    std::filesystem::path configFilePath;
    std::map<std::string, std::string> configData;
};

#endif // MODEL_MANAGER_H
