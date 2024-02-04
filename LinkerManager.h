#pragma once
#include <filesystem>
#include <map>
#include <fstream>
#include <sstream>
#include <string>

class LinkConfig {
public:
    std::string name; //文件名（显示名称，不包含后缀）
    int id; //自定义的ID，从10000累加
    std::string path; //文件路径
};

class LinkerManager {
public:
    LinkerManager() {
        int idCounter = 17750;

        // 获取当前进程的路径
        std::filesystem::path currentPath = GetBasePath();
        // 创建AutoLinker目录
        std::filesystem::path autoLinkerPath = currentPath / "AutoLinker" / "Config";
        if (!std::filesystem::exists(autoLinkerPath)) {
            std::filesystem::create_directory(autoLinkerPath);
        }

        for (const auto& entry : std::filesystem::directory_iterator(autoLinkerPath)) {
            if (entry.is_regular_file() && entry.path().extension() == ".ini") {
                LinkConfig config;
                config.name = entry.path().stem().string();
                config.id = idCounter++;
                config.path = entry.path().string();

                configs_[config.name] = config;
            }
        }
    }

    const LinkConfig& getConfig(const std::string& name) const {
        // 如果map中没有对应的键，std::map::at会抛出一个std::out_of_range异常。
        // 你可能想要在这里添加一些错误处理代码。

        if (configs_.contains(name)) {
            return configs_.at(name);
        }

        return e;
    }

    const LinkConfig& getConfig(int id) const {
        // 如果map中没有对应的键，std::map::at会抛出一个std::out_of_range异常。
        // 你可能想要在这里添加一些错误处理代码。

        for (const auto& [k, v] : configs_) {
            if (v.id == id) {
                return v;
            }
        }
        return e;
    }


    const auto getCount() const {
        return configs_.size();
    }

    const auto getMap() const {
        return configs_;
    }

private:
    std::map<std::string, LinkConfig> configs_;

    LinkConfig e;
};