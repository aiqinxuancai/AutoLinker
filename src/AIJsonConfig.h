// AIJsonConfig.h
// 管理 AIConfig.json，专门存储 AI 相关配置（key、prompt、模型参数等）。
// 优先于 INI 文件，支持从 INI 自动迁移。
// 内部始终以 UTF-8 存储，提供自动本地编码转换接口。
#pragma once

#ifndef AI_JSON_CONFIG_H
#define AI_JSON_CONFIG_H

#include <filesystem>
#include <map>
#include <string>

class AIJsonConfig {
public:
    AIJsonConfig();

    // 获取配置值（UTF-8 编码），键不存在则返回空字符串。
    std::string getValue(const std::string& key) const;

    // 获取配置值并转换为本地编码（ANSI/GBK），供 MFC 对话框直接使用。
    std::string getValueLocal(const std::string& key) const;

    // 设置配置值（本地编码，自动转 UTF-8 后持久化）。
    void setValue(const std::string& key, const std::string& localValue);

    // 批量设置多个配置值（本地编码），只写盘一次。
    void setValues(const std::map<std::string, std::string>& localPairs);

    // 是否含有任何配置数据。
    bool hasAnyData() const;

    // 是否存在指定键。
    bool hasKey(const std::string& key) const;

private:
    void load();
    void save() const;

    std::filesystem::path m_filePath;
    // 内部值均为 UTF-8 编码
    std::map<std::string, std::string> m_data;
};

#endif // AI_JSON_CONFIG_H
