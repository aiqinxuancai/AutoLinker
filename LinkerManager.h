#pragma once
#include <filesystem>
#include <map>
#include <fstream>
#include <sstream>
#include <string>

class LinkConfig {
public:
    std::string name; //�ļ�������ʾ���ƣ���������׺��
    int id; //�Զ����ID����10000�ۼ�
    std::string path; //�ļ�·��
};

class LinkerManager {
public:
    LinkerManager() {
        int idCounter = 17750;

        // ��ȡ��ǰ���̵�·��
        std::filesystem::path currentPath = GetBasePath();
        // ����AutoLinkerĿ¼
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
        // ���map��û�ж�Ӧ�ļ���std::map::at���׳�һ��std::out_of_range�쳣��
        // �������Ҫ���������һЩ��������롣

        if (configs_.contains(name)) {
            return configs_.at(name);
        }

        return e;
    }

    const LinkConfig& getConfig(int id) const {
        // ���map��û�ж�Ӧ�ļ���std::map::at���׳�һ��std::out_of_range�쳣��
        // �������Ҫ���������һЩ��������롣

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